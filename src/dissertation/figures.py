"""Рисунки по ГОСТ 7.32-2017 (п. 6.5) / ГОСТ 2.105.

* Рисунок — отдельный абзац «ГОСТ Рисунок» по центру, «не отрывать от следующего»;
* подпись под ним по центру: «Рисунок 1 – Название», без точки в конце;
* нумерация сквозная (или в пределах главы: «Рисунок 2.3»); в приложениях — с буквой
  приложения: «Рисунок А.1»;
* ширина — не больше ширины текста, высота — не больше ~2/3 страницы, пропорции сохраняются.

Номер в подписи лежит в закладке ``_GostFig_F3`` → ссылка в тексте ``{@F3}`` («на рисунке
{@F3}» → «на рисунке 2») обновляется при добавлении/перестановке рисунков.

В окне ввода рисунок записывается маркером на отдельной строке:
    ⟦Рисунок: Название | C:\\путь\\к\\файлу.png⟧           (размер — автоматически)
    ⟦Рисунок: Название | C:\\путь\\к\\файлу.png | 12 см⟧   (ширина 12 см)
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass

from docx.oxml.ns import qn
from docx.shared import Cm
from docx.text.paragraph import Paragraph

from . import fields, ooxml
from .outline import Kind, StyleResolver, classify_paragraph
from .runs import make_run, visible_text
from .style_manager import Role
from .xref import FIG_PREFIX, next_key

NEW_FIG_RX = re.compile(r"⟦\s*Рисунок\s*:\s*(?P<title>[^|⟧]*?)\s*\|\s*(?P<path>[^|⟧]+?)\s*"
                        r"(?:\|\s*(?P<width>\d+(?:[.,]\d+)?)\s*см\s*)?⟧")
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff")


@dataclass
class FigureSpec:
    path: str
    title: str
    width_cm: float | None = None

    def token(self) -> str:
        w = f" | {self.width_cm:g} см" if self.width_cm else ""
        return f"⟦Рисунок: {self.title} | {self.path}{w}⟧"


def parse_token(m: re.Match) -> FigureSpec:
    w = m.group("width")
    return FigureSpec(path=m.group("path").strip(), title=clean_title(m.group("title")),
                      width_cm=float(w.replace(",", ".")) if w else None)


def clean_title(title: str) -> str:
    """ГОСТ: с прописной буквы, без точки в конце."""
    t = " ".join(title.split()).rstrip(".").strip()
    return t[:1].upper() + t[1:] if t else t


def split_figure_tokens(text: str) -> list:
    """Текст → [str | FigureSpec] (маркеры рисунков отделены от текста)."""
    out, pos = [], 0
    for m in NEW_FIG_RX.finditer(text):
        if m.start() > pos:
            out.append(text[pos:m.start()])
        out.append(parse_token(m))
        pos = m.end()
    if pos < len(text):
        out.append(text[pos:])
    return out


def _size(builder, path: str, width_cm: float | None) -> Cm:
    from docx.image.image import Image

    pg = builder.profile.page
    # 97% ширины текста: при центрировании «впритык» LibreOffice сдвигает рисунок на поле
    max_w = (pg.width_mm - pg.margin_left_mm - pg.margin_right_mm) / 10 * 0.97
    max_h = (pg.height_mm - pg.margin_top_mm - pg.margin_bottom_mm) / 10 * 0.65
    img = Image.from_file(path)
    px_w, px_h = img.px_width, img.px_height
    dpi = img.horz_dpi or 96
    natural = px_w / dpi * 2.54
    w = min(width_cm or natural, max_w)
    if px_w and w * px_h / px_w > max_h:
        w = max_h * px_w / px_h
    return Cm(max(w, 1.0))


def insert_figure_after(builder, cursor, spec: FigureSpec):
    """Абзац с рисунком + подпись после элемента cursor. Возвращает абзац подписи."""
    if not os.path.isfile(spec.path):
        raise ValueError(f"Файл изображения не найден: {spec.path}")
    body = builder.document._body
    p_img = ooxml.new_paragraph_after(cursor, body)
    p_img.style = builder.styles.style_for(Role.FIGURE)
    p_img.add_run().add_picture(spec.path, width=_size(builder, spec.path, spec.width_cm))
    p_cap = ooxml.new_paragraph_after(p_img._p, body)
    p_cap.style = builder.styles.style_for(Role.FIGURE_CAPTION)
    key = next_key(builder.document, "F", FIG_PREFIX)        # ключ сразу — на рисунок можно сослаться
    _build_caption(builder, p_cap._p, key, "?", spec.title or os.path.splitext(os.path.basename(spec.path))[0])
    return p_cap._p


def _build_caption(builder, p_el, key: str | None, number: str, title: str, bm_id: int | None = None):
    for child in list(p_el):
        if child.tag != qn("w:pPr"):
            p_el.remove(child)
    fmt = builder.profile.figure_caption_format
    before, _, rest = fmt.partition("{n}")
    middle, _, after = rest.partition("{title}")
    if before:
        p_el.append(make_run(before))
    num = make_run(number)
    p_el.append(num)
    if key is not None:
        fields.wrap_in_bookmark(num, num, FIG_PREFIX + key,
                                bm_id if bm_id is not None else fields.next_bookmark_id(builder.document))
    p_el.append(make_run(middle + clean_title(title) + after))


@dataclass
class FigureEntry:
    key: str | None
    caption: object          # абзац подписи
    image: object | None     # абзац с рисунком (перед подписью)
    number: str
    title: str


class Figures:
    def __init__(self, builder):
        self.builder = builder
        self.document = builder.document
        self.profile = builder.profile

    @staticmethod
    def key_of(p_el) -> str | None:
        for b in p_el.iter(fields.W_BM_START):
            name = b.get(qn("w:name")) or ""
            if name.startswith(FIG_PREFIX):
                return name[len(FIG_PREFIX):]
        return None

    def title_of(self, cap_el) -> str:
        text = visible_text(Paragraph(cap_el, self.document._body))
        fmt = self.profile.figure_caption_format
        before, _, rest = fmt.partition("{n}")
        middle = rest.partition("{title}")[0]
        rx = re.compile("^" + re.escape(before.strip()) + r"\s*[\dА-ЯA-Z.?]+" + re.escape(middle.rstrip())
                        + r"\s*")
        return rx.sub("", text, count=1).strip()

    def entries(self) -> list[FigureEntry]:
        resolver = StyleResolver(self.document)
        out = []
        chapter = 0
        app_index = -1
        per = 0
        mode = self.profile.figure_numbering
        for el in ooxml.iter_body_blocks(self.document.element.body):
            if el.tag != ooxml.W_P:
                continue
            item = classify_paragraph(el, resolver)
            if item is not None and item.level == 0:
                if item.kind == Kind.CHAPTER:
                    chapter += 1
                    if mode == "chapter":
                        per = 0
                elif item.kind == Kind.APPENDIX:
                    app_index += 1
                    per = 0
                elif app_index >= 0:
                    app_index = -1
            if resolver.role_of(el) != Role.FIGURE_CAPTION:
                continue
            per += 1
            if app_index >= 0:
                number = f"{self.profile.appendix_letter(app_index)}.{per}"
            elif mode == "chapter" and chapter:
                number = f"{chapter}.{per}"
            else:
                number = str(per)
            prev = el.getprevious()
            image = prev if prev is not None and prev.tag == ooxml.W_P and \
                resolver.role_of(prev) == Role.FIGURE else None
            out.append(FigureEntry(self.key_of(el), el, image, number, self.title_of(el)))
        return out

    def numbers(self) -> dict[str, str]:
        return {e.key: e.number for e in self.entries() if e.key}

    def refresh(self) -> None:
        bm_id = fields.next_bookmark_id(self.document)
        taken: set[str] = set()
        for e in self.entries():
            if e.key is None:
                key = next_key(self.document, "F", FIG_PREFIX, taken)
                _build_caption(self.builder, e.caption, key, e.number, e.title, bm_id)
                bm_id += 1
                taken.add(key)
                continue
            taken.add(e.key)
            start = fields.find_bookmark(e.caption, FIG_PREFIX + e.key)
            run = start.getnext() if start is not None else None
            t = run.find(qn("w:t")) if run is not None and run.tag == qn("w:r") else None
            if t is None:
                _build_caption(self.builder, e.caption, e.key, e.number, e.title, bm_id)
                bm_id += 1
            elif t.text != e.number:
                t.text = e.number

    def set_title(self, cap_el, title: str) -> None:
        key = self.key_of(cap_el)
        number = next((e.number for e in self.entries() if e.caption is cap_el), "?")
        _build_caption(self.builder, cap_el, key, number, title)
