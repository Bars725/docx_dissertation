"""Оглавление/содержание — поле Word ``TOC \\o "1-3" \\h \\z \\u`` в блоке docPart.

* Word: при открытии предлагает обновить поля (флаг updateFields) и строит
  оглавление с номерами страниц; кнопка «Обновить таблицу» тоже работает.
* LibreOffice: Сервис → Обновить → Все указатели.
* Программа сама переписывает кэш (пункты без номеров страниц) при каждом
  изменении структуры — оглавление всегда соответствует заголовкам.
"""
from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from . import fields, ooxml
from .outline import Kind
from .runs import make_run
from .style_manager import Role

GALLERY = "Table of Contents"
SKIP_KINDS = {Kind.TITLE, Kind.TOC}


def find_toc_sdt(document):
    for sdt in document.element.body.iter(qn("w:sdt")):
        gallery = sdt.find(".//" + qn("w:docPartGallery"))
        if gallery is not None and gallery.get(qn("w:val")) == GALLERY:
            return sdt
    return None


def _instr(levels: int) -> str:
    return f'TOC \\o "1-{levels}" \\h \\z \\u'


def toc_entries(builder) -> list[tuple[int, str]]:
    levels = builder.profile.toc_levels
    return [(it.level, it.text) for it in builder.outline()
            if it.kind not in SKIP_KINDS and it.level < levels and it.text]


def _fill(builder, content) -> None:
    for child in list(content):
        content.remove(child)
    styles = builder.styles.ensure_toc_styles()
    entries = toc_entries(builder) or [(0, "Оглавление обновится при открытии документа в Word (F9)")]
    runs_begin = fields.field_runs(_instr(builder.profile.toc_levels), "")
    begin, instr, sep, _result, end = runs_begin
    for i, (level, text) in enumerate(entries):
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ps = OxmlElement("w:pStyle")
        ps.set(qn("w:val"), styles[min(level, 2)].style_id)
        ppr.append(ps)
        p.append(ppr)
        if i == 0:
            p.extend([begin, instr, sep])
        p.append(make_run(text + "\t"))
        if i == len(entries) - 1:
            p.append(end)
        content.append(p)


def insert_toc(builder) -> list[str]:
    """Вставить оглавление после титульного листа (или обновить существующее)."""
    warnings = []
    sdt = find_toc_sdt(builder.document)
    if sdt is not None:
        refresh_toc(builder)
        return ["Оглавление уже есть — обновлено"]
    outline = builder.outline()
    anchor = next((it.element for it in outline if it.kind != Kind.TITLE), None)
    add = builder._inserter("before", anchor) if anchor is not None else builder._inserter("end", None)
    title = builder._structural_text(builder.profile.toc_title)
    head = add(Role.TOC_HEADING, title)

    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    obj = OxmlElement("w:docPartObj")
    gal = OxmlElement("w:docPartGallery")
    gal.set(qn("w:val"), GALLERY)
    obj.append(gal)
    obj.append(OxmlElement("w:docPartUnique"))
    pr.append(obj)
    sdt.append(pr)
    content = OxmlElement("w:sdtContent")
    sdt.append(content)
    head._p.addnext(sdt)
    _fill(builder, content)
    fields.enable_update_fields_on_open(builder.document)
    if outline and outline[0].kind == Kind.TITLE:
        builder._break_before_if_needed(head._p)
    return warnings


def refresh_toc(builder) -> bool:
    sdt = find_toc_sdt(builder.document)
    if sdt is None:
        return False
    content = sdt.find(qn("w:sdtContent"))
    _fill(builder, content)
    return True


def has_toc(builder) -> bool:
    return find_toc_sdt(builder.document) is not None


__all__ = ["insert_toc", "refresh_toc", "has_toc", "find_toc_sdt", "toc_entries", "ooxml"]
