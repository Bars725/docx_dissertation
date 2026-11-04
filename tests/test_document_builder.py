import copy

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from dissertation import ooxml
from dissertation.document_builder import BlockType as B, DocumentBuilder, InsertPosition
from dissertation.gost import gost_7_32_2017
from dissertation.outline import Kind


def styles_and_texts(path):
    return [(p.style.name, p.text) for p in Document(str(path)).paragraphs]


@pytest.fixture
def builder():
    return DocumentBuilder.new(gost_7_32_2017())


def test_new_document_page_and_styles(builder, tmp_path):
    path = builder.save(tmp_path / "new.docx")
    doc = Document(str(path))
    s = doc.sections[0]
    assert round(s.page_width.mm) == 210 and round(s.page_height.mm) == 297  # не Letter!
    assert round(s.left_margin.mm) == 30 and round(s.right_margin.mm) == 15
    for name in ("ГОСТ Основной текст", "Heading 1", "Heading 2", "ГОСТ Список литературы"):
        rpr = doc.styles[name].element.rPr
        rfonts = rpr.find(qn("w:rFonts"))
        for slot in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            assert rfonts.get(qn(slot)) == "Times New Roman", (name, slot)
        for theme in ("w:asciiTheme", "w:hAnsiTheme"):
            assert rfonts.get(qn(theme)) is None
        color = rpr.find(qn("w:color"))
        assert color.get(qn("w:val")) == "000000" and color.get(qn("w:themeColor")) is None
    h1 = doc.styles["Heading 1"].paragraph_format
    assert h1.page_break_before and h1.keep_with_next
    assert ooxml.has_page_field(s.footer._element)
    assert s.different_first_page_header_footer
    assert not ooxml.has_page_field(s.first_page_footer._element)


def test_schema_order_in_ppr(builder):
    """Дочерние элементы pPr в стилях — строго в порядке схемы."""
    order = ooxml.PPR_SEQ
    for st in builder.document.styles.element.findall(qn("w:style")):
        ppr = st.find(qn("w:pPr"))
        if ppr is None:
            continue
        tags = [etree.QName(c).localname for c in ppr]
        idx = [order.index("w:" + t) for t in tags]
        assert idx == sorted(idx), (st.get(qn("w:styleId")), tags)


def test_auto_position_keeps_gost_structure(builder, tmp_path):
    builder.insert_block(B.APPENDIX, "Данные", heading="Таблицы")
    builder.insert_block(B.BIBLIOGRAPHY, "Иванов И.И. Книга.")
    builder.insert_block(B.CONCLUSION, "Выводы.")
    builder.insert_block(B.CHAPTER, "Текст главы.", heading="Теория")
    builder.insert_block(B.INTRODUCTION, "Актуальность.")
    builder.insert_block(B.TITLE_PAGE, "ДИССЕРТАЦИЯ\n\nМосква 2026")
    kinds = [it.kind for it in builder.outline_for_display()]
    assert kinds == [Kind.TITLE, Kind.INTRODUCTION, Kind.CHAPTER, Kind.CONCLUSION,
                     Kind.BIBLIOGRAPHY, Kind.APPENDIX]
    path = builder.save(tmp_path / "order.docx")
    texts = [t for _, t in styles_and_texts(path)]
    assert texts[:3] == ["ДИССЕРТАЦИЯ", "", "Москва 2026"]
    assert "1 Теория" in texts and "1. Иванов И.И. Книга." in texts and "ПРИЛОЖЕНИЕ А" in texts


def test_title_page_followed_by_page_break(builder):
    builder.insert_block(B.BODY, "Просто текст.")
    builder.insert_block(B.TITLE_PAGE, "Титул")
    paras = builder.document.paragraphs
    assert paras[0].text == "Титул"
    assert paras[1].paragraph_format.page_break_before


def test_renumber_after_insert_in_middle(builder):
    builder.insert_block(B.CHAPTER, "a", heading="Первая")
    builder.insert_block(B.SUBCHAPTER, "a", heading="Первая.один")
    builder.insert_block(B.CHAPTER, "b", heading="Вторая")
    builder.insert_block(B.SUBCHAPTER, "b", heading="Вторая.один")
    first = next(it for it in builder.outline() if it.text.endswith("Первая"))
    builder.insert_block(B.CHAPTER, "new", heading="Вставленная",
                         position=InsertPosition("section_end", first.element))
    heads = [it.text for it in builder.outline() if it.kind in (Kind.CHAPTER, Kind.SUBCHAPTER)]
    assert heads == ["1 Первая", "1.1 Первая.один", "2 Вставленная", "3 Вторая", "3.1 Вторая.один"]


def test_manual_numbers_and_dots_stripped(builder):
    builder.insert_block(B.CHAPTER, "", heading="Глава 7. Обзор.")
    builder.insert_block(B.SUBCHAPTER, "", heading="7.3 Методы")
    heads = [it.text for it in builder.outline()]
    assert heads == ["1 Обзор", "1.1 Методы"]


def test_heading_from_first_line_when_empty(builder):
    res = builder.insert_block(B.CHAPTER, "Название главы\nТекст главы.")
    assert [p.text for p in res.paragraphs] == ["1 Название главы", "Текст главы."]
    assert res.warnings


def test_bibliography_appends_and_continues_numbering(builder):
    builder.insert_block(B.BIBLIOGRAPHY, "1. А\n2. Б")
    builder.insert_block(B.APPENDIX, "x", heading="Прил")
    res = builder.insert_block(B.BIBLIOGRAPHY, "[7] В", position=InsertPosition("end"))
    assert res.warnings  # список уже есть — добавлено в его конец
    bib = [p.text for p in builder.document.paragraphs if p.style.name == "ГОСТ Список литературы"]
    assert bib == ["1. А", "2. Б", "3. В"]
    assert builder.document.paragraphs[-1].text == "x"  # приложение осталось последним


def test_appendix_letters_renumbered(builder):
    builder.insert_block(B.APPENDIX, "1", heading="Первое")
    builder.insert_block(B.APPENDIX, "2", heading="Второе")
    first = next(it for it in builder.outline() if it.kind == Kind.APPENDIX)
    builder.insert_block(B.APPENDIX, "0", heading="Новое", position=InsertPosition("after", first.element))
    labels = [it.text for it in builder.outline() if it.kind == Kind.APPENDIX]
    assert labels == ["ПРИЛОЖЕНИЕ А", "ПРИЛОЖЕНИЕ Б", "ПРИЛОЖЕНИЕ В"]


def test_inserted_runs_have_no_direct_formatting(builder):
    res = builder.insert_block(B.CHAPTER, "Абзац один.\n\nАбзац два.", heading="Тест")
    for p in res.paragraphs:
        for r in p.runs:
            assert r._r.find(qn("w:rPr")) is None
        ppr = p._p.find(qn("w:pPr"))
        assert [c.tag for c in ppr] == [qn("w:pStyle")]


def test_open_existing_keeps_its_styles_and_format(word_like_template, tmp_path):
    before = copy.deepcopy(Document(str(word_like_template)).styles["Heading 1"].element)
    b = DocumentBuilder.open(word_like_template)
    b.insert_block(B.CHAPTER, "Новый текст третьей главы.", heading="Результаты")
    path = b.save(tmp_path / "continued.docx")
    doc = Document(str(path))
    after = doc.styles["Heading 1"].element
    c14n = dict(method="c14n", exclusive=True)
    assert etree.tostring(before, **c14n) == etree.tostring(after, **c14n)  # чужие стили не тронуты
    texts = [p.text for p in doc.paragraphs]
    # формат номера взят из документа; глава встала перед списком литературы
    assert texts.index("Глава 3. Результаты") < texts.index("СПИСОК ЛИТЕРАТУРЫ")
    assert doc.styles["ГОСТ Основной текст"].font.name == "Arial"
    b2 = DocumentBuilder.open(path)
    b2.insert_block(B.BIBLIOGRAPHY, "Сидоров С.С. Новая книга.")
    assert b2.document.paragraphs[-1].text == "[3] Сидоров С.С. Новая книга."


def test_open_with_enforced_profile(word_like_template):
    b = DocumentBuilder.open(word_like_template, profile=gost_7_32_2017(), enforce_styles=True)
    assert round(b.document.sections[0].left_margin.mm) == 30
    assert b.document.styles["Heading 1"].font.size.pt == 14


def test_save_is_atomic_with_backup(builder, tmp_path):
    target = tmp_path / "doc.docx"
    builder.insert_block(B.BODY, "v1")
    builder.save(target)
    builder.insert_block(B.BODY, "v2")
    builder.save(target)
    assert (tmp_path / "doc.bak.docx").exists()
    assert [p.text for p in Document(str(tmp_path / "doc.bak.docx")).paragraphs] == ["v1"]
    assert [p.text for p in Document(str(target)).paragraphs] == ["v1", "v2"]
    assert not list(tmp_path.glob("*.~tmp.docx"))


def test_empty_insert_rejected(builder):
    with pytest.raises(ValueError):
        builder.insert_block(B.BODY, "   \n ")


def test_controller_flow(tmp_path, word_like_template):
    from dissertation.controller import AppController

    ctl = AppController(profiles_dir=tmp_path / "profiles")
    prof, saved = ctl.load_template(word_like_template)
    assert saved.exists()
    keys = [k for k, _ in ctl.available_profiles()]
    assert f"file:{saved}" in keys
    ctl.new_document(f"file:{saved}")
    ctl.insert(B.CHAPTER, "Текст", heading="Глава из образца")
    out = ctl.save(tmp_path / "from_profile.docx")
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert texts[0] == "Глава 1. Глава из образца"
    assert ctl.preview(B.BIBLIOGRAPHY, "1. A\n2) B") == ["A", "B"]
