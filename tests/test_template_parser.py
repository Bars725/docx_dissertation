from docx import Document

from dissertation.template_parser import TemplateParser
from dissertation.document_builder import DocumentBuilder
from dissertation.gost import gost_r_7_0_11_2011


def test_extracts_word_like_template(word_like_template):
    prof = TemplateParser.from_path(word_like_template).extract()
    pg = prof.page
    assert (pg.margin_left_mm, pg.margin_right_mm, pg.margin_top_mm, pg.margin_bottom_mm) == (25, 10, 20, 20)

    b = prof.body
    assert b.font_name == "Arial"
    assert b.font_size_pt == 13
    assert b.line_spacing == 2.0
    assert b.first_line_indent_cm == 1.0
    assert b.alignment == "justify"

    ch = prof.chapter_heading
    assert ch.font_size_pt == 16 and ch.bold and ch.alignment == "center"
    assert ch.page_break_before
    assert prof.subchapter_heading.font_size_pt == 14

    assert prof.numbering.enabled
    assert prof.numbering.chapter == "Глава {n}. {title}"
    assert prof.numbering.subchapter == "{n}.{m}. {title}"

    assert prof.bibliography_title == "СПИСОК ЛИТЕРАТУРЫ"
    assert prof.bibliography_number_format == "[{n}] "

    pn = prof.page_numbering
    assert pn.enabled and pn.position == "header" and pn.alignment == "center"
    assert pn.hide_on_first_page
    assert prof.structural_uppercase
    assert any("Основной текст" in line for line in prof.extraction_report)


def test_theme_fonts_resolved():
    """Шрифт не задан явно — берётся из темы (в шаблоне python-docx minor=Cambria, major=Calibri)."""
    doc = Document()
    doc.add_heading("ВВЕДЕНИЕ", 1)
    doc.add_paragraph("Текст без явного шрифта, достаточно длинный, чтобы считаться основным текстом документа.")
    prof = TemplateParser(doc).extract()
    assert prof.body.font_name == "Cambria"
    assert prof.structural_heading.font_name == "Calibri"


def test_roundtrip_builder_to_parser(tmp_path):
    """Профиль → документ → профиль: ключевые параметры совпадают."""
    src = gost_r_7_0_11_2011()
    b = DocumentBuilder.new(src)
    b.insert_block("introduction", "Актуальность темы исследования определяется рядом причин, "
                                   "описанных далее в тексте работы подробно.")
    b.insert_block("chapter", "Содержание главы, достаточно длинное, чтобы попасть в статистику "
                              "основного текста документа.", heading="Обзор")
    b.insert_block("subchapter", "Текст.", heading="Подраздел")
    path = b.save(tmp_path / "rt.docx")
    got = TemplateParser.from_path(path).extract()
    assert got.page == src.page
    for attr in ("font_name", "font_size_pt", "line_spacing", "first_line_indent_cm", "alignment"):
        assert getattr(got.body, attr) == getattr(src.body, attr), attr
    assert got.chapter_heading.alignment == "center"
    assert got.numbering.chapter == "Глава {n}. {title}"
    assert got.page_numbering.position == "header"


def test_document_without_anything_falls_back_to_gost():
    prof = TemplateParser(Document()).extract()
    assert prof.chapter_heading.bold
    assert any("не найден" in line for line in prof.extraction_report)
