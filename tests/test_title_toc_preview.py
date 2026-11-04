from docx import Document
from docx.oxml.ns import qn

from dissertation import toc
from dissertation.document_builder import BlockType as B, DocumentBuilder, InsertPosition
from dissertation.gost import gost_7_32_2017, gost_r_7_0_11_2011
from dissertation.outline import Kind
from dissertation.style_manager import Role
from dissertation.title_page import TitlePageData


def form(**kw):
    base = dict(author="Иванов Иван Иванович", title="Тема исследования",
                specialty="2.3.5. Математическое обеспечение",
                supervisor="д-р техн. наук, проф.\nПетров Пётр Петрович")
    base.update(kw)
    return TitlePageData(**base)


def test_title_form_roundtrip(tmp_path):
    b = DocumentBuilder.new(gost_r_7_0_11_2011())
    b.insert_block(B.INTRODUCTION, "Текст.")
    b.set_title_page(form())
    path = b.save(tmp_path / "t.docx")
    b2 = DocumentBuilder.open(path)
    data = b2.title_data()
    assert data.author == "Иванов Иван Иванович"
    assert data.supervisor == "д-р техн. наук, проф.\nПетров Пётр Петрович"
    assert data.manuscript is True
    kinds = [it.kind for it in b2.outline_for_display()]
    assert kinds[:2] == [Kind.TITLE, Kind.INTRODUCTION]
    # следующий за титулом элемент начинается с новой страницы
    assert b2.document.paragraphs[0].text == "ВВЕДЕНИЕ"


def test_title_form_replaces_previous(tmp_path):
    b = DocumentBuilder.new(gost_7_32_2017())
    b.insert_block(B.TITLE_PAGE, "Старый\nтитул")           # «текстовый» титул
    b.insert_block(B.BODY, "Текст.")
    b.set_title_page(form(manuscript=False))
    b.set_title_page(form(author="Другой Автор"))
    assert len(b.document.tables) == 1
    assert b.title_data().author == "Другой Автор"
    assert all(p.text not in ("Старый", "титул") for p in b.document.paragraphs)


def test_title_rows_fit_page():
    b = DocumentBuilder.new(gost_7_32_2017())
    b.set_title_page(form())
    tbl = b.document.tables[0]
    total_mm = sum(r.height.mm for r in tbl.rows)
    pg = b.profile.page
    assert total_mm < pg.height_mm - pg.margin_top_mm - pg.margin_bottom_mm


def test_toc_inserted_after_title_and_refreshed(tmp_path):
    b = DocumentBuilder.new(gost_7_32_2017())
    b.set_title_page(form())
    b.insert_block(B.INTRODUCTION, "Текст.")
    b.insert_block(B.TOC)
    b.insert_block(B.CHAPTER, "Текст.", heading="Обзор")
    kinds = [it.kind for it in b.outline_for_display()]
    assert kinds[:3] == [Kind.TITLE, Kind.TOC, Kind.INTRODUCTION]
    entries = [t for _, t in toc.toc_entries(b)]
    assert entries == ["ВВЕДЕНИЕ", "1 Обзор"]             # СОДЕРЖАНИЕ сам в оглавление не входит
    path = b.save(tmp_path / "toc.docx")
    doc = Document(str(path))
    sdt = toc.find_toc_sdt(doc)
    instr = "".join(t.text for t in sdt.iter(qn("w:instrText")))
    assert instr.strip().startswith('TOC \\o "1-3"')
    cached = ["".join(t.text for t in p.iter(qn("w:t"))).strip() for p in sdt.iter(qn("w:p"))]
    assert cached == ["ВВЕДЕНИЕ", "1 Обзор"]
    assert doc.settings.element.find(qn("w:updateFields")) is not None
    res = DocumentBuilder.open(path).insert_block(B.TOC)
    assert res.warnings                                   # повторно — только обновление


def test_preview_does_not_modify_and_shows_numbers():
    b = DocumentBuilder.new(gost_7_32_2017())
    b.insert_block(B.CHAPTER, "a", heading="Первая")
    b.insert_block(B.CONCLUSION, "Выводы.")
    b.insert_block(B.BIBLIOGRAPHY, "Иванов. Книга.")
    before = len(b.document.element.body)
    pr = b.preview_block(B.CHAPTER, "Текст [@R1].", heading="Вторая")
    assert len(b.document.element.body) == before
    assert pr.items == [(Role.CHAPTER, "2 Вторая"), (Role.BODY, "Текст [1].")]
    assert pr.page_breaks == [True, False]
    assert pr.where == "перед «ЗАКЛЮЧЕНИЕ»"
    first = b.outline()[0]
    pr2 = b.preview_block(B.SUBCHAPTER, "x", heading="Пара",
                          position=InsertPosition("section_end", first.element))
    assert pr2.items[0] == (Role.SUBCHAPTER, "1.1 Пара")
    pr3 = b.preview_block(B.BIBLIOGRAPHY, "Петров. Статья.")
    assert [t for _, t in pr3.items] == ["1. Иванов. Книга.", "2. Петров. Статья."]


def test_preview_toc_and_appendix():
    b = DocumentBuilder.new(gost_7_32_2017())
    b.insert_block(B.APPENDIX, "x", heading="Первое")
    assert b.preview_block(B.APPENDIX, "", heading="Второе").items[0] == (Role.APPENDIX, "ПРИЛОЖЕНИЕ Б")
    assert b.preview_block(B.TOC).items[0] == (Role.TOC_HEADING, "СОДЕРЖАНИЕ")
