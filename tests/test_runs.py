from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from dissertation import ooxml
from dissertation.runs import fill_placeholders, replace_text, set_paragraph_text, visible_text


def test_replace_across_split_runs_keeps_first_run_format(placeholder_doc):
    doc = Document(str(placeholder_doc))
    p = doc.paragraphs[0]
    assert replace_text(p, "{{TITLE}}", "Оформление по ГОСТ") == 1
    assert p.text == "Тема: Оформление по ГОСТ — конец"
    assert p.runs[0].bold is True
    assert p.runs[0].text == "Тема: Оформление по ГОСТ"


def test_fill_placeholders_everywhere(placeholder_doc, tmp_path):
    doc = Document(str(placeholder_doc))
    n = fill_placeholders(ooxml.iter_all_paragraphs(doc),
                          {"{{TITLE}}": "Т", "{{AUTHOR}}": "Автор", "{{CITY}}": "Москва"})
    assert n == 3  # 4-й абзац разрезан табуляцией — не трогаем
    out = tmp_path / "filled.docx"
    doc.save(str(out))
    doc2 = Document(str(out))
    assert doc2.tables[0].cell(0, 0).text == "Автор"
    assert doc2.sections[0].header.paragraphs[0].text == "Город Москва"
    assert "{{TI" in doc2.paragraphs[1].text


def test_multiple_occurrences_and_count():
    doc = Document()
    p = doc.add_paragraph()
    for piece in ("a-b", "-a", "-b-a"):
        p.add_run(piece)
    assert replace_text(p, "a", "XX") == 3
    assert p.text == "XX-b-XX-b-XX"
    p2 = doc.add_paragraph("1 1 1")
    assert replace_text(p2, "1", "2", count=1) == 1
    assert p2.text == "2 1 1"


def test_anchored_start_only():
    doc = Document()
    p = doc.add_paragraph("  1 Глава 1")
    assert replace_text(p, "1 ", "2 ", anchored_start=True) == 1
    assert p.text == "  2 Глава 1"


def test_set_paragraph_text_keeps_paragraph_props_and_bookmarks():
    doc = Document()
    p = doc.add_paragraph("старый ", style="Heading 1")
    p.add_run("текст").italic = True
    bm = OxmlElement("w:bookmarkStart")
    bm.set(qn("w:id"), "0")
    bm.set(qn("w:name"), "_Toc1")
    p._p.insert(1, bm)
    p.paragraph_format.keep_with_next = True
    set_paragraph_text(p, "новый\tтекст")
    assert p.style.name == "Heading 1"
    assert p.paragraph_format.keep_with_next
    assert p._p.find(qn("w:bookmarkStart")) is not None
    assert visible_text(p) == "новыйтекст"
    assert p.text == "новый\tтекст"
    assert len(p.runs) == 1


def test_field_not_damaged():
    doc = Document()
    p = doc.add_paragraph("Стр. ")
    ooxml.append_field(p, "PAGE")
    replace_text(p, "Стр. 1", "X")   # совпадение пересекает fldChar-барьер — замены нет
    assert p.text.startswith("Стр. ")
    assert ooxml.has_page_field(p._p)
    assert p._p.xpath("count(.//w:fldChar)") == 3
