import pytest
from docx import Document

from dissertation import fields
from dissertation.bibliography import alpha_key, parse_citation
from dissertation.document_builder import BlockType as B, DocumentBuilder
from dissertation.gost import gost_7_32_2017


@pytest.fixture
def b():
    builder = DocumentBuilder.new(gost_7_32_2017())
    builder.insert_block(B.BIBLIOGRAPHY, "Сидоров С. С. Книга. – 2020.\n"
                                         "Abrams A. Paper. – 2019.\n"
                                         "Андреев А. А. Статья. – 2021.")
    return builder


def texts(builder):
    return [p.text for p in builder.document.paragraphs]


def test_parse_citation():
    assert [(p.key, p.locator) for p in parse_citation("R3, с. 15; @R5")] == [("R3", "с. 15"), ("R5", "")]
    assert parse_citation("не ссылка") is None


def test_entries_get_keys_and_numbers(b):
    assert [(e.key, e.text) for e in b.bibliography.entries()] == [
        ("R1", "Сидоров С. С. Книга. – 2020."), ("R2", "Abrams A. Paper. – 2019."),
        ("R3", "Андреев А. А. Статья. – 2021.")]
    assert "1. Сидоров С. С. Книга. – 2020." in texts(b)


def test_citations_follow_sorting_and_insertion(b, tmp_path):
    b.insert_block(B.CHAPTER, "Как в [@R3, с. 15] и [@R1; @R2].", heading="Обзор")
    assert "Как в [3, с. 15] и [1; 2]." in texts(b)
    b.bibliography.sort_alphabetically()   # кириллица раньше латиницы
    assert [e.key for e in b.bibliography.entries()] == ["R3", "R1", "R2"]
    assert "Как в [1, с. 15] и [2; 3]." in texts(b)
    b.bibliography.add(["Аарон А. Самая первая. – 2000."])
    b.bibliography.sort_alphabetically()
    assert "Как в [2, с. 15] и [3; 4]." in texts(b)
    path = b.save(tmp_path / "refs.docx")
    doc = Document(str(path))
    instrs = [f.instr.strip() for f in fields.iter_complex_fields(doc.element.body)]
    assert 'REF _GostRef_R3 \\h' in instrs           # «живые» поля для Word (F9)
    assert fields.find_bookmark(doc.element.body, "_GostRef_R3") is not None


def test_sort_by_citation_order(b):
    b.insert_block(B.CHAPTER, "Сначала [@R2], потом [@R3].", heading="Глава")
    b.bibliography.sort_by_citation()
    assert [e.key for e in b.bibliography.entries()] == ["R2", "R3", "R1"]
    assert "Сначала [1], потом [2]." in texts(b)


def test_move_update_remove(b):
    b.insert_block(B.CHAPTER, "См. [@R2].", heading="Глава")
    b.bibliography.move("R2", -1)
    assert [e.key for e in b.bibliography.entries()] == ["R2", "R1", "R3"]
    assert "См. [1]." in texts(b)
    b.bibliography.update_text("R1", "Сидоров С. С. Книга. 2-е изд. – 2022.")
    assert "2. Сидоров С. С. Книга. 2-е изд. – 2022." in texts(b)
    assert b.bibliography.remove("R2") == 1           # на запись была 1 ссылка
    assert "См. [?]." in texts(b)


def test_unknown_reference_warns(b):
    res = b.insert_block(B.BODY, "Ссылка [@R42].")
    assert any("R42" in w for w in res.warnings)
    assert "Ссылка [?]." in texts(b)


def test_alpha_key_cyrillic_first():
    items = ["Zeta", "Яковлев", "alpha", "Ёлкин", "Андреев"]
    assert sorted(items, key=alpha_key) == ["Андреев", "Ёлкин", "Яковлев", "alpha", "Zeta"]


def test_keys_not_reused_while_cited(b):
    b.insert_block(B.BODY, "Ссылка [@R3].")
    b.bibliography.remove("R3")
    added = b.bibliography.add(["Новый источник."])
    assert added[0].key == "R4"
    assert "Ссылка [?]." in texts(b)          # старая ссылка не «переехала» на новый источник


def test_legacy_entries_adopted(word_like_template):
    builder = DocumentBuilder.open(word_like_template)
    entries = builder.bibliography.entries()
    assert [e.key for e in entries] == [None, None]
    assert entries[0].text.startswith("Иванов И. И.")
    builder.bibliography.set_entries(entries)       # «принять» записи под управление
    assert [e.key for e in builder.bibliography.entries()] == ["R1", "R2"]
    assert texts(builder)[-1].startswith("[2] Петров")
