"""Титульный лист как форма (ГОСТ Р 7.0.11-2011, приложение А).

Разметка — невидимая таблица 1×4 со строками точной высоты: верх (организация,
«На правах рукописи», автор), середина (название, специальность, степень),
научный руководитель и низ (город – год). Это надёжно держит «город и год
внизу страницы» и в Word, и в LibreOffice — в отличие от пустых строк.

Каждое значение формы обёрнуто скрытой закладкой ``_GostTitle_<поле>``, поэтому
при повторном открытии документа форма заполняется из него, а повторное
сохранение формы заменяет титульный лист целиком.
"""
from __future__ import annotations

from dataclasses import dataclass, fields as dc_fields
from datetime import date

from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

from . import fields, ooxml
from .outline import TITLE_BOOKMARK_PREFIX, Kind
from .runs import make_run
from .style_manager import ALIGN_MAP, Role

V_ALIGN = {"top": WD_CELL_VERTICAL_ALIGNMENT.TOP, "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
           "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}


@dataclass
class TitlePageData:
    organization: str = ("Министерство науки и высшего образования Российской Федерации\n"
                         "Федеральное государственное бюджетное образовательное учреждение\n"
                         "высшего образования «Название университета»")
    manuscript: bool = True
    author: str = ""
    title: str = ""
    specialty: str = ""
    degree: str = "Диссертация на соискание учёной степени\nкандидата  наук"
    supervisor: str = ""
    consultant: str = ""
    city: str = "Москва"
    year: str = str(date.today().year)


# (поле, подпись в форме, многострочное)
FORM_FIELDS = [
    ("organization", "Организация", True),
    ("manuscript", "«На правах рукописи»", False),
    ("author", "ФИО автора", False),
    ("title", "Тема диссертации", True),
    ("specialty", "Специальность (шифр и название)", False),
    ("degree", "Соискание степени", True),
    ("supervisor", "Научный руководитель (степень, звание, ФИО)", True),
    ("consultant", "Научный консультант (если есть)", True),
    ("city", "Город", False),
    ("year", "Год", False),
]


@dataclass
class _Line:
    key: str | None              # поле формы или None (только литерал)
    align: str = "center"
    bold: bool = False
    caps: bool = False
    before_pt: float = 0
    prefix: str = ""
    left_indent_cm: float = 0
    suffix_key: str | None = None   # второе поле в той же строке (город – год)
    joiner: str = ""


def _layout(data: TitlePageData, half_cm: float):
    sup_lines = []
    if data.supervisor.strip():
        sup_lines += [_Line(None, "left", prefix="Научный руководитель:", left_indent_cm=half_cm),
                      _Line("supervisor", "left", left_indent_cm=half_cm)]
    if data.consultant.strip():
        sup_lines += [_Line(None, "left", prefix="Научный консультант:", left_indent_cm=half_cm,
                            before_pt=12),
                      _Line("consultant", "left", left_indent_cm=half_cm)]
    top = [_Line("organization", "center")]
    if data.manuscript:
        top.append(_Line("manuscript", "right", before_pt=36, prefix="На правах рукописи"))
    top.append(_Line("author", "center", before_pt=36))
    return [
        (0.30, "top", top),
        (0.38, "center", [_Line("title", "center", bold=True, caps=True),
                          _Line("specialty", "center", before_pt=24, prefix="Специальность "),
                          _Line("degree", "center", before_pt=24)]),
        (0.24, "center", sup_lines or [_Line(None)]),
        (0.08, "bottom", [_Line("city", "center", suffix_key="year", joiner=" – ")]),
    ]


def preview_lines(data: TitlePageData) -> list[tuple[str, str, bool, int]]:
    """Для предпросмотра формы: [(текст, выравнивание, полужирный, отступ_сверху_строк)]."""
    out = []
    for r, (_share, _valign, lines) in enumerate(_layout(data, 0)):
        for i, line in enumerate(lines):
            if line.key == "manuscript":
                text = line.prefix
            else:
                parts = [line.prefix] if line.prefix else []
                if line.key:
                    parts.append(str(getattr(data, line.key)).strip())
                if line.suffix_key:
                    parts.append(line.joiner + str(getattr(data, line.suffix_key)).strip())
                text = "".join(parts)
            if line.caps:
                text = text.upper()
            align = "right" if line.align == "right" or line.left_indent_cm else line.align
            gap = (3 if r else 0) if i == 0 else (1 if line.before_pt else 0)
            out.append((text, align, line.bold, gap))
    return out


def read_title_data(document) -> TitlePageData | None:
    body = document.element.body
    names = fields.bookmark_names(body, TITLE_BOOKMARK_PREFIX)
    if not names:
        return None
    data = TitlePageData()
    for f in dc_fields(TitlePageData):
        if f.name == "manuscript":
            data.manuscript = (TITLE_BOOKMARK_PREFIX + "manuscript") in names
            continue
        value = fields.bookmark_text(body, TITLE_BOOKMARK_PREFIX + f.name)
        setattr(data, f.name, value if value is not None else "")
    return data


def _no_borders(tbl) -> None:
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)
    mar = OxmlElement("w:tblCellMar")
    for side in ("left", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def set_title_page(builder, data: TitlePageData):
    """Создать или заменить титульный лист по данным формы. Возвращает элемент таблицы."""
    doc = builder.document
    body_el = doc.element.body
    old = [it.element for it in builder.outline() if it.kind == Kind.TITLE]

    pg = builder.profile.page
    width_cm = (pg.width_mm - pg.margin_left_mm - pg.margin_right_mm) / 10
    usable_mm = pg.height_mm - pg.margin_top_mm - pg.margin_bottom_mm - 10
    title_style = builder.styles.style_for(Role.TITLE)

    tbl = doc.add_table(rows=4, cols=1)
    tbl.autofit = False
    _no_borders(tbl)
    bm_id = fields.next_bookmark_id(doc)
    for (share, valign, lines), row in zip(_layout(data, width_cm / 2), tbl.rows):
        row.height = Mm(usable_mm * share)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        cell = row.cells[0]
        cell.width = Cm(width_cm)
        cell.vertical_alignment = V_ALIGN[valign]
        for i, line in enumerate(lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.style = title_style
            pf = p.paragraph_format
            pf.alignment = ALIGN_MAP[line.align]
            if line.before_pt:
                pf.space_before = Pt(line.before_pt)
            if line.left_indent_cm:
                pf.left_indent = Cm(line.left_indent_cm)
            bm_id = _fill_line(p, line, data, bm_id)

    # на место старого титула (или в начало документа)
    anchor = old[0] if old else next(
        (c for c in body_el if c.tag != ooxml.W_SECTPR and c is not tbl._tbl), None)
    if anchor is not None:
        anchor.addprevious(tbl._tbl)
    for el in old:
        el.getparent().remove(el)
    builder._break_after(tbl._tbl)
    return tbl._tbl


def _fill_line(p, line: _Line, data: TitlePageData, bm_id: int) -> int:
    p_el = p._p

    def styled(run):
        rpr = OxmlElement("w:rPr")
        if line.bold:
            rpr.append(OxmlElement("w:b"))
        if line.caps:
            rpr.append(OxmlElement("w:caps"))
        if len(rpr):
            run.insert(0, rpr)
        return run

    if line.key == "manuscript":
        r = styled(make_run(line.prefix))
        p_el.append(r)
        fields.wrap_in_bookmark(r, r, TITLE_BOOKMARK_PREFIX + "manuscript", bm_id)
        return bm_id + 1
    if line.prefix:
        p_el.append(styled(make_run(line.prefix)))
    for key, joiner in ((line.key, ""), (line.suffix_key, line.joiner)):
        if key is None:
            continue
        if joiner:
            p_el.append(styled(make_run(joiner)))
        r = styled(make_run(str(getattr(data, key)).strip()))
        p_el.append(r)
        fields.wrap_in_bookmark(r, r, TITLE_BOOKMARK_PREFIX + key, bm_id)
        bm_id += 1
    return bm_id
