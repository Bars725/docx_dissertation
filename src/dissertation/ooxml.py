"""Низкоуровневые операции над WordprocessingML, которых нет в python-docx.

Главное правило модуля: дочерние элементы w:pPr / w:rPr вставляются строго в
порядке схемы OOXML. Word молча «чинит» или отказывается открывать файл,
если порядок нарушен, а LibreOffice может проигнорировать свойство.
"""
from __future__ import annotations

from typing import Iterable, Iterator

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# Порядок дочерних элементов по схеме (CT_PPrBase + хвост CT_PPr).
PPR_SEQ = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
)

RPR_SEQ = (
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
    "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
    "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
    "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
    "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
)

FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia")
THEME_FONT_ATTRS = ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme")

W_P = qn("w:p")
W_TBL = qn("w:tbl")
W_SDT = qn("w:sdt")
W_SECTPR = qn("w:sectPr")


# --------------------------------------------------------------- порядок схемы
def get_or_insert(parent, tag: str, seq: tuple[str, ...]):
    """Найти дочерний элемент или создать его на правильной позиции по схеме."""
    existing = parent.find(qn(tag))
    if existing is not None:
        return existing
    el = OxmlElement(tag)
    idx = seq.index(tag)
    successors = {qn(t) for t in seq[idx + 1:]}
    for child in parent:
        if child.tag in successors:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def remove_child(parent, tag: str) -> None:
    for el in parent.findall(qn(tag)):
        parent.remove(el)


def set_on_off(parent, tag: str, seq: tuple[str, ...], value: bool) -> None:
    """Булево свойство (w:b, w:caps, w:keepNext...): явно true/false, чтобы перебить наследование."""
    el = get_or_insert(parent, tag, seq)
    if value:
        el.attrib.pop(qn("w:val"), None)
    else:
        el.set(qn("w:val"), "0")


def read_on_off(el) -> bool | None:
    """Значение булева свойства или None, если элемента нет."""
    if el is None:
        return None
    val = el.get(qn("w:val"))
    return val not in ("0", "false", "off")


# ---------------------------------------------------------------- шрифты/язык
def set_fonts(rpr, font_name: str) -> None:
    """Явно задать шрифт для всех четырёх слотов и убрать ссылки на шрифты темы.

    Ловушка python-docx: ``font.name = ...`` пишет только w:ascii/w:hAnsi.
    Если в стиле осталось w:asciiTheme="minorHAnsi" — Word возьмёт шрифт темы
    (Calibri), а кириллица/«сложные» скрипты (w:cs, w:eastAsia) останутся
    прежними. Это и есть «ломающиеся шрифты в колонтитулах».
    """
    rfonts = get_or_insert(rpr, "w:rFonts", RPR_SEQ)
    for attr in THEME_FONT_ATTRS:
        rfonts.attrib.pop(qn(attr), None)
    for attr in FONT_ATTRS:
        rfonts.set(qn(attr), font_name)


def set_size(rpr, size_pt: float) -> None:
    half_points = str(int(round(size_pt * 2)))
    get_or_insert(rpr, "w:sz", RPR_SEQ).set(qn("w:val"), half_points)
    get_or_insert(rpr, "w:szCs", RPR_SEQ).set(qn("w:val"), half_points)


def set_black_color(rpr) -> None:
    """Чёрный цвет без ссылки на цвет темы (стандартные Heading в Word синие)."""
    color = get_or_insert(rpr, "w:color", RPR_SEQ)
    for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
        color.attrib.pop(qn(attr), None)
    color.set(qn("w:val"), "000000")


def set_lang(rpr, lang: str) -> None:
    el = get_or_insert(rpr, "w:lang", RPR_SEQ)
    el.set(qn("w:val"), lang)
    el.set(qn("w:eastAsia"), lang)


def get_doc_defaults(styles_element):
    """(rPr, pPr) из w:docDefaults, создаются при отсутствии."""
    dd = styles_element.find(qn("w:docDefaults"))
    if dd is None:
        dd = OxmlElement("w:docDefaults")
        styles_element.insert(0, dd)
    rpr_default = dd.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = OxmlElement("w:rPrDefault")
        dd.insert(0, rpr_default)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rpr_default.append(rpr)
    ppr_default = dd.find(qn("w:pPrDefault"))
    if ppr_default is None:
        ppr_default = OxmlElement("w:pPrDefault")
        rpr_default.addnext(ppr_default)
    ppr = ppr_default.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        ppr_default.append(ppr)
    return rpr, ppr


# ---------------------------------------------------------------------- поля
def append_field(paragraph: Paragraph, instr: str, placeholder: str = "1") -> None:
    """Сложное поле (fldChar begin/separate/end) — надёжнее w:fldSimple в Word и LibreOffice."""
    p = paragraph._p

    def run_with(child):
        r = OxmlElement("w:r")
        r.append(child)
        p.append(r)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run_with(begin)

    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    run_with(instr_el)

    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    run_with(sep)

    t = OxmlElement("w:t")
    t.text = placeholder
    run_with(t)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_with(end)


def has_page_field(element) -> bool:
    """Есть ли в части документа поле PAGE (сложное или простое)."""
    for instr in element.iter(qn("w:instrText")):
        if instr.text and instr.text.strip().split(" ")[0].upper() == "PAGE":
            return True
    for fld in element.iter(qn("w:fldSimple")):
        if (fld.get(qn("w:instr")) or "").strip().upper().startswith("PAGE"):
            return True
    return False


# ------------------------------------------------------------ обход документа
def iter_body_blocks(body_element) -> Iterator:
    """Абзацы и таблицы верхнего уровня тела (w:sdt раскрываются)."""
    for child in body_element.iterchildren():
        if child.tag in (W_P, W_TBL):
            yield child
        elif child.tag == W_SDT:
            content = child.find(qn("w:sdtContent"))
            if content is not None:
                yield from iter_body_blocks(content)


def iter_all_paragraph_elements(document) -> Iterator:
    """Все w:p документа: тело, таблицы (вложенные), надписи, колонтитулы всех секций."""
    seen_parts = set()
    yield from document.element.body.iter(W_P)
    for section in document.sections:
        for hf in (section.header, section.first_page_header, section.even_page_header,
                   section.footer, section.first_page_footer, section.even_page_footer):
            if hf.is_linked_to_previous:
                continue
            el = hf._element
            if id(el) in seen_parts:
                continue
            seen_parts.add(id(el))
            yield from el.iter(W_P)


def iter_all_paragraphs(document) -> Iterator[Paragraph]:
    for p in iter_all_paragraph_elements(document):
        yield Paragraph(p, document._body)


def new_paragraph_after(anchor_element, parent) -> Paragraph:
    """Создать пустой абзац сразу после элемента (абзаца/таблицы)."""
    p = OxmlElement("w:p")
    anchor_element.addnext(p)
    return Paragraph(p, parent)


def new_paragraph_before(anchor_element, parent) -> Paragraph:
    p = OxmlElement("w:p")
    anchor_element.addprevious(p)
    return Paragraph(p, parent)


def contains_page_break(element) -> bool:
    for br in element.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def paragraph_has_section_break(p_element) -> bool:
    ppr = p_element.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:sectPr")) is not None


def element_text(element) -> str:
    return "".join(t.text or "" for t in element.iter(qn("w:t")))


def first(iterable: Iterable, default=None):
    return next(iter(iterable), default)
