"""Поля Word (REF, TOC), закладки и settings.xml.

Поля — способ сделать «живые» части документа, которые пересчитывает сам Word
(F9, печать, открытие с updateFields): ссылки на номер источника и оглавление.
Программа при каждом сохранении тоже обновляет их кэшированные значения,
поэтому документ выглядит правильно и в LibreOffice / просмотрщиках.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterator

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_R = qn("w:r")
W_T = qn("w:t")
W_FLDCHAR = qn("w:fldChar")
W_INSTR = qn("w:instrText")
W_BM_START = qn("w:bookmarkStart")
W_BM_END = qn("w:bookmarkEnd")

# Порядок дочерних элементов w:settings (фрагмент схемы, нужный нам).
SETTINGS_AFTER_UPDATE_FIELDS = (
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars", "w:rsids",
    "m:mathPr", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary",
    "o:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
)
_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
}


def _clark(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{_NS[prefix]}}}{local}"


# ================================================================ сложные поля
@dataclass
class ComplexField:
    instr: str = ""
    result_t: list = field(default_factory=list)   # w:t в результате поля
    begin_run: object = None
    end_run: object = None


def iter_complex_fields(root) -> Iterator[ComplexField]:
    """Все сложные поля внутри элемента (в порядке документа, вложенные — раньше внешних)."""
    stack: list[tuple[ComplexField, str]] = []
    for r in root.iter(W_R):
        for child in r:
            if child.tag == W_FLDCHAR:
                kind = child.get(qn("w:fldCharType"))
                if kind == "begin":
                    f = ComplexField(begin_run=r)
                    stack.append((f, "instr"))
                elif kind == "separate" and stack:
                    stack[-1] = (stack[-1][0], "result")
                elif kind == "end" and stack:
                    f, _ = stack.pop()
                    f.end_run = r
                    yield f
            elif child.tag == W_INSTR and stack and stack[-1][1] == "instr":
                stack[-1][0].instr += child.text or ""
            elif child.tag == W_T and stack and stack[-1][1] == "result":
                stack[-1][0].result_t.append(child)


def set_field_result(f: ComplexField, value: str) -> None:
    if not f.result_t:
        return
    f.result_t[0].text = value
    for t in f.result_t[1:]:
        t.text = ""


def field_runs(instr: str, result: str) -> list:
    """Runs сложного поля: begin, instrText, separate, результат, end."""
    runs = []

    def run_with(child):
        r = OxmlElement("w:r")
        r.append(child)
        runs.append(r)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run_with(begin)
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    run_with(it)
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    run_with(sep)
    t = OxmlElement("w:t")
    t.text = result
    run_with(t)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_with(end)
    return runs


REF_RX = re.compile(r"^\s*REF\s+(\S+)", re.I)


def ref_target(instr: str) -> str | None:
    m = REF_RX.match(instr)
    return m.group(1) if m else None


# ==================================================================== закладки
def next_bookmark_id(document) -> int:
    ids = [int(b.get(qn("w:id"), "0")) for b in document.element.body.iter(W_BM_START)
           if (b.get(qn("w:id")) or "").isdigit()]
    return max(ids, default=0) + 1


def wrap_in_bookmark(first_el, last_el, name: str, bm_id: int) -> None:
    """Обернуть последовательность соседних элементов (runs) закладкой."""
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bm_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bm_id))
    first_el.addprevious(start)
    last_el.addnext(end)


def find_bookmark(root, name: str):
    for b in root.iter(W_BM_START):
        if b.get(qn("w:name")) == name:
            return b
    return None


def bookmark_names(root, prefix: str = "") -> list[str]:
    return [b.get(qn("w:name")) for b in root.iter(W_BM_START)
            if (b.get(qn("w:name")) or "").startswith(prefix)]


def bookmark_text(root, name: str) -> str | None:
    """Текст между bookmarkStart и bookmarkEnd (w:br → перевод строки)."""
    start = find_bookmark(root, name)
    if start is None:
        return None
    bm_id = start.get(qn("w:id"))
    out, inside, para_seen = [], False, None
    for el in root.iter():
        if el is start:
            inside = True
            continue
        if not inside:
            continue
        if el.tag == W_BM_END and el.get(qn("w:id")) == bm_id:
            break
        if el.tag == qn("w:p"):
            if para_seen is not None:
                out.append("\n")
            para_seen = el
        elif el.tag == W_T:
            out.append(el.text or "")
        elif el.tag in (qn("w:br"), qn("w:cr")):
            out.append("\n")
        elif el.tag == qn("w:tab"):
            out.append("\t")
    return "".join(out)


# ==================================================================== settings
def set_setting_flag(document, tag: str, successors: tuple[str, ...]) -> None:
    """Добавить в settings.xml флаг (w:updateFields, ...) на правильную позицию."""
    settings = document.settings.element
    if settings.find(_clark(tag)) is not None:
        return
    el = OxmlElement(tag)
    el.set(qn("w:val"), "true")
    succ = {_clark(t) for t in successors}
    for child in settings:
        if child.tag in succ:
            child.addprevious(el)
            return
    settings.append(el)


def enable_update_fields_on_open(document) -> None:
    """Word при открытии предложит обновить поля (оглавление, ссылки) — с номерами страниц."""
    set_setting_flag(document, "w:updateFields", SETTINGS_AFTER_UPDATE_FIELDS)
