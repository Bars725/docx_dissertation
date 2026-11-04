"""Совместимость: документ пересохраняется LibreOffice и дорабатывается снова.

Запускается только при наличии soffice. Аналогично стоит прогонять файлы,
пересохранённые в Word и экспортированные из Google Docs (см. docs/TESTING.md).
"""
import pytest
from docx import Document

from conftest import libreoffice_convert, soffice
from dissertation import ooxml
from dissertation.document_builder import BlockType as B, DocumentBuilder
from dissertation.gost import gost_7_32_2017
from dissertation.outline import Kind
from dissertation.template_parser import TemplateParser

pytestmark = [pytest.mark.libreoffice,
              pytest.mark.skipif(soffice() is None, reason="LibreOffice не установлен")]


@pytest.fixture
def built(tmp_path):
    b = DocumentBuilder.new(gost_7_32_2017())
    b.insert_block(B.TITLE_PAGE, "ДИССЕРТАЦИЯ\n\nМосква 2026")
    b.insert_block(B.INTRODUCTION, "Актуальность темы исследования определяется рядом причин, "
                                   "которые подробно рассмотрены далее в тексте работы.")
    b.insert_block(B.CHAPTER, "Текст первой главы достаточно длинный для статистики основного "
                              "текста документа и проверки параметров.", heading="Обзор")
    b.insert_block(B.SUBCHAPTER, "Текст.", heading="Методы")
    b.insert_block(B.BIBLIOGRAPHY, "Иванов И.И. Книга.\nПетров П.П. Статья.")
    b.insert_block(B.APPENDIX, "Данные.", heading="Таблицы")
    (tmp_path / "src").mkdir()
    return b.save(tmp_path / "src" / "built.docx")


def test_libreoffice_resave_keeps_structure_and_format(built, tmp_path):
    resaved = libreoffice_convert(built, "docx:MS Word 2007 XML", tmp_path / "lo")
    doc = Document(str(resaved))
    names = {p.style.name for p in doc.paragraphs}
    assert {"ГОСТ Основной текст", "ГОСТ Список литературы"} <= names
    prof = TemplateParser(doc, source=str(resaved)).extract()
    assert prof.body.font_name == "Times New Roman" and prof.body.font_size_pt == 14
    assert prof.body.line_spacing == 1.5 and prof.body.first_line_indent_cm == 1.25
    assert round(prof.page.margin_left_mm) == 30
    assert prof.page_numbering.enabled and prof.page_numbering.hide_on_first_page

    # Дорабатываем пересохранённый файл — нумерация продолжается корректно.
    b = DocumentBuilder.open(resaved)
    b.insert_block(B.CHAPTER, "Вторая глава.", heading="Эксперимент")
    b.insert_block(B.BIBLIOGRAPHY, "Сидоров С.С. Ещё.")
    kinds = [it.kind for it in b.outline_for_display()]
    assert kinds.index(Kind.CHAPTER) < kinds.index(Kind.BIBLIOGRAPHY) < kinds.index(Kind.APPENDIX)
    texts = [p.text for p in b.document.paragraphs]
    assert "2 Эксперимент" in texts and "3. Сидоров С.С. Ещё." in texts
    out = b.save(tmp_path / "again.docx")
    assert ooxml.has_page_field(Document(str(out)).sections[0].footer._element)


def test_pdf_render_page_count(built, tmp_path):
    """Титул, введение, глава, список, приложение — каждый с новой страницы (подглава — нет)."""
    import re

    pdf = libreoffice_convert(built, "pdf", tmp_path / "pdf")
    pages = len(re.findall(rb"/Type\s*/Page(?!s)", pdf.read_bytes()))
    assert pages == 5
