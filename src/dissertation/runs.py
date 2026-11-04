"""Операции над текстом абзаца, устойчивые к дроблению на runs.

Word режет текст абзаца на произвольные w:r: из-за проверки орфографии
(w:proofErr), истории правок (rsid), смены языка ввода, автозамены.
Плейсхолдер ``{{TITLE}}`` в XML легко оказывается тремя runs: ``{{TI``,
``TLE``, ``}}``. Поэтому мы никогда не ищем текст «внутри run», а работаем
с абзацем как с единой строкой и отображаем её обратно на элементы w:t.
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

W_T = qn("w:t")
MC_ALT = "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
W_R = qn("w:r")
# Элементы, которые разрывают «склейку» текста: совпадение не может их пересечь.
BARRIERS = {qn("w:tab"), qn("w:br"), qn("w:cr"), qn("w:fldChar"), qn("w:drawing"),
            qn("w:pict"), qn("w:object"), qn("w:footnoteReference"),
            qn("w:endnoteReference"), qn("w:sym")}
# Не спускаемся во вложенные «миры» с собственными абзацами (надписи, фигуры).
OPAQUE = {qn("w:txbxContent"), MC_ALT, qn("w:drawing"), qn("w:pict"),
          qn("w:object"), qn("w:del"), qn("w:moveFrom")}
SENTINEL = "\x00"

# Содержимое абзаца, которое удаляется при полной замене текста.
# pPr, закладки и диапазоны комментариев сохраняем.
_CONTENT_TAGS = {qn("w:r"), qn("w:hyperlink"), qn("w:fldSimple"), qn("w:smartTag"),
                 qn("w:proofErr"), qn("w:ins"), qn("w:del"), qn("w:customXml"),
                 qn("w:sdt"), qn("w:moveTo"), qn("w:moveFrom")}


@dataclass
class _Segment:
    element: object | None   # w:t или None для барьера
    start: int
    length: int


def _walk(el):
    for child in el:
        tag = child.tag
        if tag in OPAQUE:
            if tag in BARRIERS or tag == MC_ALT:
                yield None
            continue
        if tag == W_T:
            yield child
        elif tag in BARRIERS:
            yield None
        else:
            yield from _walk(child)


def _segments(p_element) -> tuple[str, list[_Segment]]:
    text_parts: list[str] = []
    segments: list[_Segment] = []
    pos = 0
    for el in _walk(p_element):
        if el is None:
            segments.append(_Segment(None, pos, 1))
            text_parts.append(SENTINEL)
            pos += 1
        else:
            t = el.text or ""
            segments.append(_Segment(el, pos, len(t)))
            text_parts.append(t)
            pos += len(t)
    return "".join(text_parts), segments


def visible_text(paragraph: Paragraph) -> str:
    """Текст абзаца, включая гиперссылки и вставки правок; барьеры → пробел/таб."""
    text, _ = _segments(paragraph._p)
    return text.replace(SENTINEL, "")


def _set_t(t_el, value: str) -> None:
    t_el.text = value
    if value != value.strip() or "  " in value:
        t_el.set(qn("xml:space"), "preserve")


def replace_text(paragraph: Paragraph, old: str, new: str, count: int = 0,
                 anchored_start: bool = False) -> int:
    """Заменить ``old`` на ``new`` с учётом раздробленных runs.

    Замена получает форматирование того run, где начинается совпадение;
    остальные затронутые w:t укорачиваются. Позиции, пересекающие барьеры
    (табуляция, разрыв строки, поле), не совпадают — это защищает поля
    (номер страницы, ссылки) от повреждения.

    count=0 — заменить все вхождения. anchored_start — только в начале абзаца
    (с учётом ведущих пробелов). Возвращает число замен.
    """
    if not old:
        return 0
    full, segments = _segments(paragraph._p)
    if anchored_start:
        lead = len(full) - len(full.lstrip(" "))
        starts = [lead] if full.startswith(old, lead) else []
    else:
        starts = [m.start() for m in re.finditer(re.escape(old), full)]
    if count:
        starts = starts[:count]
    # С конца — тогда смещения более ранних совпадений остаются валидными.
    for s in reversed(starts):
        e = s + len(old)
        first = True
        for seg in segments:
            if seg.element is None or seg.length == 0:
                continue
            seg_end = seg.start + seg.length
            if seg_end <= s or seg.start >= e:
                continue
            txt = seg.element.text or ""
            ls = max(s, seg.start) - seg.start
            le = min(e, seg_end) - seg.start
            if first:
                _set_t(seg.element, txt[:ls] + new + txt[le:])
                first = False
            else:
                _set_t(seg.element, txt[:ls] + txt[le:])
    if starts:
        _drop_empty_runs(paragraph._p)
    return len(starts)


def _drop_empty_runs(p_element) -> None:
    """Удалить runs, в которых после замены не осталось ничего, кроме пустых w:t и rPr."""
    for r in list(p_element.iter(W_R)):
        meaningful = [c for c in r if c.tag != qn("w:rPr")]
        if meaningful and all(c.tag == W_T and not (c.text or "") for c in meaningful):
            r.getparent().remove(r)


def set_paragraph_text(paragraph: Paragraph, text: str, keep_run_format: bool = True) -> None:
    """Полностью заменить текст абзаца, сохранив свойства абзаца (стиль, отступы, разрывы).

    Форматирование символов берётся из первого run с текстом (keep_run_format)
    или не задаётся вовсе — тогда действует стиль абзаца.
    """
    p = paragraph._p
    template_rpr = None
    if keep_run_format:
        for r in p.iter(W_R):
            if r.find(W_T) is not None and (r.findtext(W_T) or ""):
                rpr = r.find(qn("w:rPr"))
                template_rpr = copy.deepcopy(rpr) if rpr is not None else None
                break
    for child in list(p):
        if child.tag in _CONTENT_TAGS:
            p.remove(child)
    p.append(make_run(text, template_rpr))


def make_run(text: str, rpr=None):
    """Run с текстом; ``\\n`` → w:br, ``\\t`` → w:tab. Без прямого форматирования, если rpr=None."""
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(rpr)
    for i, line in enumerate(text.split("\n")):
        if i:
            r.append(OxmlElement("w:br"))
        for j, chunk in enumerate(line.split("\t")):
            if j:
                r.append(OxmlElement("w:tab"))
            if chunk:
                t = OxmlElement("w:t")
                _set_t(t, chunk)
                r.append(t)
    return r


def prepend_text(paragraph: Paragraph, prefix: str) -> None:
    """Добавить текст в начало абзаца с форматированием первого текстового run."""
    _, segments = _segments(paragraph._p)
    for seg in segments:
        if seg.element is not None:
            _set_t(seg.element, prefix + (seg.element.text or ""))
            return
    paragraph._p.append(make_run(prefix))


def fill_placeholders(paragraphs, mapping: dict[str, str]) -> int:
    """Заменить плейсхолдеры (``{{KEY}}`` и т.п.) во всех переданных абзацах."""
    total = 0
    for p in paragraphs:
        text = visible_text(p)
        for key, value in mapping.items():
            if key in text:
                total += replace_text(p, key, value)
    return total
