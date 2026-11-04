import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from dissertation.document_builder import BlockType as B, DocumentBuilder, InsertPosition
from dissertation.gost import gost_7_32_2017
from dissertation.outline import Kind
from dissertation.structure_ops import StructureOps


@pytest.fixture
def b():
    builder = DocumentBuilder.new(gost_7_32_2017())
    builder.insert_block(B.INTRODUCTION, "Введение.")
    builder.insert_block(B.CHAPTER, "Текст главы один.", heading="Анализ")
    builder.insert_block(B.SUBCHAPTER, "Текст 1.1.", heading="Обзор")
    builder.insert_block(B.SECTION, "Текст А.\n\nВторой абзац А.", heading="Виды документов")
    builder.insert_block(B.SECTION, "Текст Б.", heading="Проблемы")
    builder.insert_block(B.CHAPTER, "Текст главы два.", heading="Разработка")
    builder.insert_block(B.BIBLIOGRAPHY, "Иванов. Книга.\nПетров. Статья.")
    return builder


def heads(b):
    return [it.text for it in b.outline() if it.kind in (Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION)]


def el(b, text):
    return next(it.element for it in b.outline() if it.text == text)


def test_read_section_content(b):
    ops = StructureOps(b)
    c = ops.read(el(b, "1.1.1 Виды документов"))
    assert c.heading == "Виды документов"
    assert c.body == "Текст А.\n\nВторой абзац А."
    assert c.body_editable and c.heading_editable
    chapter = ops.read(el(b, "1 Анализ"))
    assert chapter.body == "Текст главы один."          # только собственный текст, без подпунктов
    intro = ops.read(el(b, "ВВЕДЕНИЕ"))
    assert intro.heading == "ВВЕДЕНИЕ" and not intro.heading_editable


def test_read_restores_citation_tokens(b):
    b.insert_block(B.BODY, "См. [@R2, с. 5] и [@R1; @R2].",
                   position=InsertPosition("section_end", el(b, "1.1.2 Проблемы")))
    c = StructureOps(b).read(el(b, "1.1.2 Проблемы"))
    assert "См. [@R2, с. 5] и [@R1; @R2]." in c.body


def test_update_heading_and_body(b):
    ops = StructureOps(b)
    res = ops.update(el(b, "1.1.1 Виды документов"), "Классы документов.",
                     "Новый текст со ссылкой [@R1].\n\nЕщё абзац.")
    assert res.ok and not res.warnings
    c = ops.read(el(b, "1.1.1 Классы документов"))
    assert c.body == "Новый текст со ссылкой [@R1].\n\nЕщё абзац."
    texts = [p.text for p in b.document.paragraphs]
    assert "Новый текст со ссылкой [1]." in texts and "Текст А." not in texts
    assert texts.index("1.1.1 Классы документов") < texts.index("1.1.2 Проблемы")


def test_move_section_within_same_level(b):
    ops = StructureOps(b)
    assert ops.move(el(b, "1.1.2 Проблемы"), -1).ok
    assert heads(b) == ["1 Анализ", "1.1 Обзор", "1.1.1 Проблемы", "1.1.2 Виды документов",
                        "2 Разработка"]
    texts = [p.text for p in b.document.paragraphs]
    assert texts.index("Текст Б.") < texts.index("1.1.2 Виды документов") < texts.index("Второй абзац А.")
    # глава переезжает вместе с подглавами и пунктами
    assert ops.move(el(b, "2 Разработка"), -1).ok
    assert heads(b) == ["1 Разработка", "2 Анализ", "2.1 Обзор", "2.1.1 Проблемы",
                        "2.1.2 Виды документов"]


def test_move_blocked_by_level_and_structure(b):
    ops = StructureOps(b)
    r = ops.move(el(b, "1 Анализ"), -1)               # выше — ВВЕДЕНИЕ
    assert not r.ok and "ВВЕДЕНИЕ" in r.message
    assert not ops.move(el(b, "2 Разработка"), 1).ok   # ниже — список литературы
    assert not ops.move(el(b, "1.1.1 Виды документов"), -1).ok   # первый в своей подглаве
    assert not ops.move(el(b, "ВВЕДЕНИЕ"), 1).ok
    assert heads(b)[0] == "1 Анализ"


def test_delete_section_with_children(b):
    ops = StructureOps(b)
    assert "подпунктов: 3" in ops.describe(el(b, "1 Анализ"))
    assert ops.delete(el(b, "1 Анализ")).ok
    assert heads(b) == ["1 Разработка"]
    assert "Текст 1.1." not in [p.text for p in b.document.paragraphs]


def test_delete_bibliography_warns_about_refs(b):
    b.insert_block(B.BODY, "См. [@R1].")
    res = StructureOps(b).delete(el(b, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"))
    assert res.warnings
    assert "См. [?]." in [p.text for p in b.document.paragraphs]


def test_complex_elements_become_markers(b):
    target = el(b, "1.1.2 Проблемы")
    body_p = StructureOps(b).own_body(target)[0]
    run = OxmlElement("w:r")
    run.append(OxmlElement("w:drawing"))
    body_p.append(run)
    ops = StructureOps(b)
    c = ops.read(target)
    assert c.body_editable and "рисунок" in c.reason
    assert c.body.startswith("⟦#1 Рисунок")
    # маркер оставили и дописали текст вокруг — элемент сохранён и стоит после нового абзаца
    res = ops.update(target, "Проблемы ручного оформления", "Новый абзац.\n\n" + c.body + "\n\nХвост.")
    assert res.ok and not res.warnings
    assert body_p.getparent() is not None
    body_els = StructureOps(b).own_body(el(b, "1.1.2 Проблемы ручного оформления"))
    assert [ooxml_text(e) for e in body_els] == ["Новый абзац.", "Текст Б.", "Хвост."]
    # маркер удалили — элемент удалён, с предупреждением
    res = ops.update(el(b, "1.1.2 Проблемы ручного оформления"), None, "Только текст.")
    assert any("Удалены" in w for w in res.warnings)
    assert body_p.getparent() is None


def test_force_edit_soft_protected(b):
    target = el(b, "1.1.2 Проблемы")
    p = StructureOps(b).own_body(target)[0]
    ppr = p.find(qn("w:pPr"))
    num = OxmlElement("w:numPr")
    ppr.append(num)                                  # «список Word»
    ops = StructureOps(b)
    c = ops.read(target)
    assert c.soft_protected == 1 and c.body.startswith("⟦#1 Список")
    forced = ops.read(target, force=True)
    assert forced.body == "Текст Б." and forced.soft_protected == 0
    res = ops.update(target, None, "Текст Б, исправленный.", force=True)
    assert res.ok
    assert StructureOps(b).read(el(b, "1.1.2 Проблемы")).body == "Текст Б, исправленный."


def test_invisible_bookmarks_do_not_block(b):
    target = el(b, "1.1.2 Проблемы")
    bm = OxmlElement("w:bookmarkStart")
    bm.set(qn("w:id"), "99")
    bm.set(qn("w:name"), "_GoBack")
    target.addnext(bm)                               # закладка на уровне тела (бывает после Word)
    c = StructureOps(b).read(target)
    assert c.body == "Текст Б." and not c.reason
    StructureOps(b).update(target, None, "Новый.")
    assert bm.getparent() is not None


def ooxml_text(e):
    from dissertation import ooxml
    return ooxml.element_text(e)
