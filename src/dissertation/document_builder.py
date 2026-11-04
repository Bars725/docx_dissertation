"""Сборка итогового документа: вставка блоков, нумерация, колонтитулы, сохранение."""
from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm
from docx.text.paragraph import Paragraph

from . import ooxml
from .outline import Kind, OutlineItem, StyleResolver, build_outline
from .profile import FormattingProfile
from .runs import make_run, prepend_text, replace_text, set_paragraph_text, visible_text
from .style_manager import ALIGN_MAP, Role, StyleManager
from .text_cleaner import CleanOptions, TextCleaner


class BlockType(str, Enum):
    TITLE_PAGE = "title_page"
    INTRODUCTION = "introduction"
    CHAPTER = "chapter"
    SUBCHAPTER = "subchapter"
    SECTION = "section"
    BODY = "body"
    CONCLUSION = "conclusion"
    BIBLIOGRAPHY = "bibliography"
    APPENDIX = "appendix"


BLOCK_LABELS = {
    BlockType.TITLE_PAGE: "Титульный лист",
    BlockType.INTRODUCTION: "Введение",
    BlockType.CHAPTER: "Глава",
    BlockType.SUBCHAPTER: "Подглава (параграф)",
    BlockType.SECTION: "Пункт",
    BlockType.BODY: "Основной текст (продолжение)",
    BlockType.CONCLUSION: "Заключение",
    BlockType.BIBLIOGRAPHY: "Список литературы",
    BlockType.APPENDIX: "Приложение",
}

# Блоки, у которых пользователь вводит заголовок.
BLOCKS_WITH_HEADING = {BlockType.CHAPTER, BlockType.SUBCHAPTER, BlockType.SECTION, BlockType.APPENDIX}

_STRUCTURAL_KIND = {BlockType.INTRODUCTION: Kind.INTRODUCTION,
                    BlockType.CONCLUSION: Kind.CONCLUSION,
                    BlockType.BIBLIOGRAPHY: Kind.BIBLIOGRAPHY}
_MANUAL_NUMBER = re.compile(r"^\s*(?:(?:глава|раздел)\s+)?\d+(?:\.\d+)*\.?\s+", re.I)


@dataclass
class InsertPosition:
    """Куда вставлять.

    auto        — по смыслу блока (глава — перед заключением/списком/приложениями и т.д.);
    start / end — начало / конец документа;
    after       — сразу после элемента ``anchor`` (w:p);
    section_end — в конец раздела, начатого заголовком ``anchor``.
    """

    mode: str = "auto"
    anchor: object | None = None


@dataclass
class InsertResult:
    paragraphs: list[Paragraph] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class SaveError(RuntimeError):
    pass


class DocumentBuilder:
    def __init__(self, document, profile: FormattingProfile, path: str | Path | None = None):
        self.document = document
        self.profile = profile
        self.path = Path(path) if path else None
        self.styles = StyleManager(document, profile)
        self._backup_done = False

    # =========================================================== создание
    @classmethod
    def new(cls, profile: FormattingProfile) -> "DocumentBuilder":
        b = cls(Document(), profile)
        b.apply_page_setup()
        b.styles.ensure_styles(overwrite=True)
        b.setup_page_numbering()
        b._apply_settings()
        return b

    @classmethod
    def open(cls, path: str | Path, profile: FormattingProfile | None = None,
             enforce_styles: bool = False) -> "DocumentBuilder":
        """Открыть существующий документ.

        profile=None — профиль извлекается из самого документа, и новые фрагменты
        оформляются «как в документе». enforce_styles=True — стили и поля
        документа приводятся к профилю (существующий текст в этих стилях тоже изменится).
        """
        from .template_parser import TemplateParser

        doc = Document(str(path))
        if profile is None:
            profile = TemplateParser(doc, source=str(path)).extract()
        b = cls(doc, profile, path)
        b.styles.ensure_styles(overwrite=enforce_styles)
        if enforce_styles:
            b.apply_page_setup()
            if profile.page_numbering.enabled and not b.has_page_numbers():
                b.setup_page_numbering()
        return b

    # ===================================================== страница/колонтитулы
    def apply_page_setup(self) -> None:
        pg = self.profile.page
        for s in self.document.sections:
            s.orientation = WD_ORIENT.PORTRAIT
            s.page_width, s.page_height = Mm(pg.width_mm), Mm(pg.height_mm)
            s.left_margin, s.right_margin = Mm(pg.margin_left_mm), Mm(pg.margin_right_mm)
            s.top_margin, s.bottom_margin = Mm(pg.margin_top_mm), Mm(pg.margin_bottom_mm)
            s.header_distance, s.footer_distance = Mm(pg.header_distance_mm), Mm(pg.footer_distance_mm)
            s.gutter = Mm(0)

    def has_page_numbers(self) -> bool:
        for s in self.document.sections:
            for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer):
                if not hf.is_linked_to_previous and ooxml.has_page_field(hf._element):
                    return True
        return False

    def setup_page_numbering(self) -> None:
        """Номер страницы полем PAGE; на первой странице (титул) — пустой колонтитул."""
        pn = self.profile.page_numbering
        if not pn.enabled or not self.document.sections:
            return
        s = self.document.sections[0]
        s.different_first_page_header_footer = pn.hide_on_first_page
        hf = s.footer if pn.position == "footer" else s.header
        hf.is_linked_to_previous = False
        p = hf.paragraphs[0] if hf.paragraphs else hf.add_paragraph()
        for child in list(p._p):
            if child.tag != qn("w:pPr"):
                p._p.remove(child)
        style = self.styles.find_style(name="Footer" if pn.position == "footer" else "Header")
        if style is not None:
            p.style = style
        p.paragraph_format.alignment = ALIGN_MAP.get(pn.alignment)
        p.paragraph_format.first_line_indent = Mm(0)
        ooxml.append_field(p, "PAGE")
        if pn.hide_on_first_page:
            first = s.first_page_footer if pn.position == "footer" else s.first_page_header
            first.is_linked_to_previous = False  # создаёт пустой колонтитул первой страницы

    def _apply_settings(self) -> None:
        if not self.profile.auto_hyphenation:
            return
        settings = self.document.settings.element
        if settings.find(qn("w:autoHyphenation")) is not None:
            return
        el = OxmlElement("w:autoHyphenation")
        anchor = settings.find(qn("w:defaultTabStop"))
        if anchor is not None:
            anchor.addnext(el)

    # ================================================================ структура
    def outline(self) -> list[OutlineItem]:
        return build_outline(self.document)

    def outline_for_display(self) -> list[OutlineItem]:
        """Структура для GUI: титульный лист схлопывается в один пункт."""
        out, title_added = [], False
        for it in self.outline():
            if it.kind == Kind.TITLE:
                if not title_added:
                    out.append(OutlineItem(it.element, Kind.TITLE, 0, "[Титульный лист]", it.style_id))
                    title_added = True
                continue
            out.append(it)
        return out

    def find(self, kind: Kind) -> OutlineItem | None:
        return next((it for it in self.outline() if it.kind == kind), None)

    # ================================================================ вставка
    def insert_block(self, block: BlockType, text: str = "", heading: str = "",
                     position: InsertPosition | None = None, clean: bool = True,
                     clean_options: CleanOptions | None = None) -> InsertResult:
        block = BlockType(block)
        position = position or InsertPosition()
        cleaner = TextCleaner(clean_options or CleanOptions())
        result = InsertResult()

        plan = self._plan(block, text, heading, clean, cleaner, position, result)
        if not plan:
            raise ValueError("Нечего вставлять: введите текст или заголовок")

        where, anchor = self._resolve_position(block, position, result)
        add = self._inserter(where, anchor)
        for role, value in plan:
            result.paragraphs.append(add(role, value))

        if block == BlockType.TITLE_PAGE:
            self._break_after(result.paragraphs[-1]._p)
        elif where == "after" and anchor is not None and self._is_title(anchor):
            self._break_before_if_needed(result.paragraphs[0]._p)
        self.renumber()
        return result

    # --------------------------------------------------------------- план
    def _plan(self, block, text, heading, clean, cleaner, position, result):
        p = self.profile
        if clean:
            body = cleaner.to_paragraphs(text)
        else:
            body = [ln.rstrip("\r") for ln in text.split("\n") if ln.strip()]
        heading = cleaner.heading(heading) if heading.strip() else ""

        if block == BlockType.TITLE_PAGE:
            if self.find(Kind.TITLE) is not None:
                result.warnings.append("Титульный лист уже есть — новые строки добавлены в начало")
            lines = cleaner.lines(text, keep_empty=True) if clean else text.split("\n")
            return [(Role.TITLE, ln) for ln in lines]

        if block == BlockType.BODY:
            return [(Role.BODY, t) for t in body]

        if block in _STRUCTURAL_KIND:
            kind = _STRUCTURAL_KIND[block]
            title = {BlockType.INTRODUCTION: p.introduction_title,
                     BlockType.CONCLUSION: p.conclusion_title,
                     BlockType.BIBLIOGRAPHY: p.bibliography_title}[block]
            title = self._structural_text(heading or title)
            role = Role.BIBLIOGRAPHY if block == BlockType.BIBLIOGRAPHY else Role.BODY
            if block == BlockType.BIBLIOGRAPHY:
                entries = cleaner.bibliography_entries(text) if clean else body
                items = [(role, p.bibliography_number_format.format(n=0) + e) for e in entries]
            else:
                items = [(role, t) for t in body]
            if self.find(kind) is not None:
                if items:
                    result.warnings.append(f"«{title}» уже есть — добавлен только текст")
                return items
            return [(Role.STRUCTURAL, title)] + items

        if block in (BlockType.CHAPTER, BlockType.SUBCHAPTER, BlockType.SECTION):
            if not heading:
                if not body:
                    return []
                heading = cleaner.heading(body.pop(0))
                result.warnings.append("Заголовок не указан — взята первая строка текста")
            role = {BlockType.CHAPTER: Role.CHAPTER, BlockType.SUBCHAPTER: Role.SUBCHAPTER,
                    BlockType.SECTION: Role.SECTION}[block]
            return [(role, self._numbered_heading(block, heading))] + [(Role.BODY, t) for t in body]

        if block == BlockType.APPENDIX:
            if not heading and body:
                heading = cleaner.heading(body.pop(0))
            label = p.appendix_label.format(letter=p.appendix_letter(0))
            items = [(Role.APPENDIX, label)]
            if heading:
                items.append((Role.APPENDIX_TITLE, heading))
            return items + [(Role.BODY, t) for t in body]
        raise ValueError(f"Неизвестный тип блока: {block}")

    def _structural_text(self, title: str) -> str:
        if self.profile.structural_uppercase:
            return title.upper()
        return title[:1].upper() + title[1:].lower()

    def _numbered_heading(self, block: BlockType, title: str) -> str:
        num = self.profile.numbering
        if not num.enabled:
            return title
        title = _MANUAL_NUMBER.sub("", title, count=1)
        fmt = {BlockType.CHAPTER: num.chapter, BlockType.SUBCHAPTER: num.subchapter,
               BlockType.SECTION: num.section}[block]
        return fmt.format(n=0, m=0, k=0, title=title)   # номера проставит renumber()

    # ------------------------------------------------------------ позиция
    def _resolve_position(self, block, position: InsertPosition, result):
        mode, anchor = position.mode, position.anchor
        if mode in ("after", "section_end") and anchor is None:
            mode = "auto"
        if block == BlockType.BIBLIOGRAPHY:
            bib = self.find(Kind.BIBLIOGRAPHY)
            if bib is not None:
                if mode != "auto":
                    result.warnings.append("Записи добавлены в конец существующего списка литературы")
                return self._section_end(bib)
        if mode == "start" or (mode == "auto" and block == BlockType.TITLE_PAGE):
            return "start", None
        if mode == "end":
            return "end", None
        if mode == "after":
            return "after", anchor
        if mode == "section_end":
            item = next((it for it in self.outline() if it.element is anchor), None)
            return self._section_end(item) if item else ("after", anchor)

        # --- auto
        outline = self.outline()

        def before_first(*kinds):
            for it in outline:
                if it.kind in kinds:
                    return "before", it.element
            return None

        kind = _STRUCTURAL_KIND.get(block)
        if kind is not None:
            existing = next((it for it in outline if it.kind == kind), None)
            if existing is not None:
                return self._section_end(existing)
        if block == BlockType.INTRODUCTION:
            return before_first(Kind.CHAPTER, Kind.CONCLUSION, Kind.BIBLIOGRAPHY,
                                Kind.APPENDIX) or ("end", None)
        if block in (BlockType.CHAPTER, BlockType.SUBCHAPTER, BlockType.SECTION, BlockType.BODY):
            return before_first(Kind.CONCLUSION, Kind.BIBLIOGRAPHY, Kind.APPENDIX) or ("end", None)
        if block == BlockType.CONCLUSION:
            return before_first(Kind.BIBLIOGRAPHY, Kind.APPENDIX) or ("end", None)
        if block == BlockType.BIBLIOGRAPHY:
            return before_first(Kind.APPENDIX) or ("end", None)
        return "end", None

    def _section_end(self, item: OutlineItem):
        """Позиция перед следующим заголовком того же или более высокого уровня."""
        resolver = StyleResolver(self.document)
        from .outline import classify_paragraph

        started = False
        for el in ooxml.iter_body_blocks(self.document.element.body):
            if el is item.element:
                started = True
                continue
            if not started or el.tag != ooxml.W_P:
                continue
            nxt = classify_paragraph(el, resolver)
            if nxt is not None and nxt.kind != Kind.TITLE and nxt.level <= item.level:
                return "before", el
        return "end", None

    def _inserter(self, where: str, anchor):
        body = self.document._body
        state = {"cursor": anchor}
        if where == "start":
            first = ooxml.first(ooxml.iter_body_blocks(self.document.element.body))
            if first is None:
                where = "end"
            else:
                where, state["cursor"] = "before", first

        def add(role: Role, text: str) -> Paragraph:
            if where == "end":
                p = self.document.add_paragraph()
            elif where == "before":
                p = ooxml.new_paragraph_before(state["cursor"], body)
            else:
                p = ooxml.new_paragraph_after(state["cursor"], body)
                state["cursor"] = p._p
            p.style = self.styles.style_for(role)
            if text:
                p._p.append(make_run(text))   # без прямого форматирования — только стиль
            return p

        return add

    # ---------------------------------------------------------- разрывы страниц
    def _is_title(self, el) -> bool:
        return StyleResolver(self.document).role_of(el) == Role.TITLE

    def _has_break_before(self, el) -> bool:
        resolver = StyleResolver(self.document)
        for ppr in resolver.ppr_chain(el):
            v = ooxml.read_on_off(ppr.find(qn("w:pageBreakBefore")))
            if v is not None:
                return v
        return False

    def _break_before_if_needed(self, el) -> None:
        if el.tag == ooxml.W_P and not self._has_break_before(el):
            Paragraph(el, self.document._body).paragraph_format.page_break_before = True

    def _break_after(self, el) -> None:
        nxt = el.getnext()
        while nxt is not None and nxt.tag not in (ooxml.W_P, ooxml.W_TBL):
            nxt = nxt.getnext()
        if nxt is not None and nxt.tag == ooxml.W_P and not self._is_title(nxt):
            self._break_before_if_needed(nxt)

    # ============================================================== нумерация
    def renumber(self) -> None:
        """Пересчитать номера глав/подглав/пунктов, буквы приложений и номера в списке литературы.

        Меняется только префикс-номер (через run-безопасную замену), форматирование
        и остальной текст заголовка не трогаются. Заголовки, не подходящие под
        формат профиля (оформлены вручную иначе), пропускаются.
        """
        outline = self.outline()
        num = self.profile.numbering
        if num.enabled:
            n = m = k = 0
            for it in outline:
                if it.kind == Kind.CHAPTER:
                    n, m, k = n + 1, 0, 0
                    fmt = num.chapter
                elif it.kind == Kind.SUBCHAPTER and n:
                    m, k = m + 1, 0
                    fmt = num.subchapter
                elif it.kind == Kind.SECTION and n and m:
                    k += 1
                    fmt = num.section
                else:
                    continue
                self._set_number_prefix(it.element, fmt, n, m, k)

        resolver = StyleResolver(self.document)
        idx = 0
        for it in outline:
            if it.kind != Kind.APPENDIX:
                continue
            if resolver.role_of(it.element) == Role.APPENDIX:
                expected = self.profile.appendix_label.format(letter=self.profile.appendix_letter(idx))
                para = Paragraph(it.element, self.document._body)
                if visible_text(para).strip() != expected:
                    set_paragraph_text(para, expected)
            idx += 1
        self._renumber_bibliography(outline)

    def _set_number_prefix(self, el, fmt: str, n: int, m: int, k: int) -> None:
        para = Paragraph(el, self.document._body)
        text = visible_text(para).lstrip()
        match = _format_regex(fmt).match(text)
        if not match:
            return
        old_prefix = text[:match.start("title")]
        new_prefix = fmt.format(n=n, m=m, k=k, title="")
        if old_prefix != new_prefix:
            replace_text(para, old_prefix, new_prefix, count=1, anchored_start=True)

    def _renumber_bibliography(self, outline) -> None:
        bib = next((it for it in outline if it.kind == Kind.BIBLIOGRAPHY), None)
        if bib is None:
            return
        fmt = self.profile.bibliography_number_format
        rx = _format_regex(fmt.rstrip() + " {title}") if "{n}" in fmt else None
        resolver = StyleResolver(self.document)
        from .outline import classify_paragraph

        started, idx = False, 0
        for el in ooxml.iter_body_blocks(self.document.element.body):
            if el is bib.element:
                started = True
                continue
            if not started or el.tag != ooxml.W_P:
                continue
            item = classify_paragraph(el, resolver)
            if item is not None and item.level == 0:
                break
            para = Paragraph(el, self.document._body)
            text = visible_text(para).lstrip()
            if not text:
                continue
            idx += 1
            if resolver.role_of(el) != Role.BIBLIOGRAPHY or rx is None:
                continue
            new_prefix = fmt.format(n=idx)
            match = rx.match(text)
            if match:
                old_prefix = text[:match.start("title")]
                if old_prefix != new_prefix:
                    replace_text(para, old_prefix, new_prefix, count=1, anchored_start=True)
            else:
                prepend_text(para, new_prefix)

    # =============================================================== сохранение
    def save(self, path: str | Path | None = None, backup: bool = True) -> Path:
        """Атомарное сохранение: временный файл → проверка открытием → замена.

        Если Word держит файл открытым (Windows), будет понятная ошибка, а
        исходный файл не пострадает.
        """
        target = Path(path) if path else self.path
        if target is None:
            raise SaveError("Не указан путь сохранения")
        target = target.with_suffix(".docx") if target.suffix.lower() != ".docx" else target
        self.renumber()
        tmp = target.with_name(target.stem + ".~tmp.docx")
        try:
            self.document.save(str(tmp))
            Document(str(tmp))   # файл должен открываться
            if backup and target.exists() and not self._backup_done:
                shutil.copy2(target, target.with_name(target.stem + ".bak.docx"))
                self._backup_done = True
            os.replace(tmp, target)
        except PermissionError as exc:
            raise SaveError(f"Файл занят другой программой (закройте его в Word): {target}") from exc
        finally:
            if tmp.exists():
                try:
                    tmp.unlink()
                except OSError:
                    pass
        self.path = target
        return target


def _format_regex(fmt: str) -> re.Pattern:
    """«Глава {n}. {title}» → ^Глава\\s+(\\d+)\\.\\s+(?P<title>.*)$"""
    pattern = "^"
    for part in re.split(r"(\{\w+\})", fmt):
        if part in ("{n}", "{m}", "{k}"):
            pattern += r"\d+"
        elif part == "{title}":
            pattern += r"(?P<title>.*)"
        elif part:
            pattern += r"\s+".join(re.escape(chunk) for chunk in part.split(" ")).replace(
                r"\s+\s+", r"\s+") if " " in part else re.escape(part)
    return re.compile(pattern + "$", re.S | re.I)
