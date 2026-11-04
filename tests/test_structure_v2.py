"""Порядок ГОСТ, перемещение через родителей, уровни, замок, добавление элементов, приложения."""
import pytest
from docx import Document

from dissertation.document_builder import BlockType as B, DocumentBuilder, InsertPosition
from dissertation.gost import gost_7_32_2017
from dissertation.outline import Kind
from dissertation.structure_ops import StructureOps


@pytest.fixture
def b():
    x = DocumentBuilder.new(gost_7_32_2017())
    x.insert_block(B.INTRODUCTION, "Введение.")
    x.insert_block(B.CHAPTER, "Глава один.", heading="Анализ")
    x.insert_block(B.SUBCHAPTER, "Т 1.1.", heading="Обзор")
    x.insert_block(B.SECTION, "Т А.", heading="Виды")
    x.insert_block(B.SECTION, "Т Б.", heading="Проблемы")
    x.insert_block(B.SUBCHAPTER, "Т 1.2.", heading="Методы")
    x.insert_block(B.SECTION, "Т В.", heading="Подход")
    x.insert_block(B.CHAPTER, "Глава два.", heading="Разработка")
    x.insert_block(B.CONCLUSION, "Выводы.")
    return x


def heads(b):
    return [it.text for it in b.outline()]


def el(b, text):
    return next(it.element for it in b.outline() if it.text == text)


def test_order_normalized_by_gost(b):
    # вставка «после главы» — введение всё равно окажется на своём месте
    b.insert_block(B.TOC)
    b.insert_block(B.BIBLIOGRAPHY, "Иванов. Книга.")
    b.insert_block(B.APPENDIX, "x", heading="Данные")
    kinds = [it.kind for it in b.outline() if it.level == 0]
    assert kinds == [Kind.TOC, Kind.INTRODUCTION, Kind.CHAPTER, Kind.CHAPTER, Kind.CONCLUSION,
                     Kind.BIBLIOGRAPHY, Kind.APPENDIX]


def test_wrong_order_in_existing_document_fixed(b, tmp_path):
    body = b.document.element.body
    intro_sec = StructureOps(b).section_elements(el(b, "ВВЕДЕНИЕ"))
    concl = StructureOps(b).section_elements(el(b, "ЗАКЛЮЧЕНИЕ"))
    cursor = concl[-1]
    for e in intro_sec:                       # испортить: введение после заключения
        cursor.addnext(e)
        cursor = e
    assert [it.kind for it in b.outline() if it.level == 0][-1] == Kind.INTRODUCTION
    assert b.normalize_order() is True
    assert [it.kind for it in b.outline() if it.level == 0][0] == Kind.INTRODUCTION
    assert b.normalize_order() is False
    assert "Введение." in [p.text for p in b.document.paragraphs][:3]
    del body


def test_explicit_position_clamped_to_main_part(b):
    b.insert_block(B.BIBLIOGRAPHY, "Иванов. Книга.")
    res = b.insert_block(B.SUBCHAPTER, "т", heading="Лишняя",
                         position=InsertPosition("after", el(b, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")))
    assert res.warnings
    texts = heads(b)
    assert texts.index("2.1 Лишняя") < texts.index("ЗАКЛЮЧЕНИЕ")


def test_move_across_parents(b):
    ops = StructureOps(b)
    assert ops.move(el(b, "1.2.1 Подход"), -1).ok        # первый в 1.2 → последний в 1.1
    assert heads(b)[2:7] == ["1.1 Обзор", "1.1.1 Виды", "1.1.2 Проблемы", "1.1.3 Подход", "1.2 Методы"]
    assert ops.move(el(b, "1.1.3 Подход"), 1).ok          # последний в 1.1 → первый в 1.2
    assert heads(b)[5:7] == ["1.2 Методы", "1.2.1 Подход"]
    assert ops.move(el(b, "1.2 Методы"), 1).ok            # подглава → первой во 2-ю главу
    assert heads(b)[5:8] == ["2 Разработка", "2.1 Методы", "2.1.1 Подход"]
    texts = [p.text for p in b.document.paragraphs]
    assert texts.index("Глава два.") < texts.index("2.1 Методы")   # после собственного текста главы
    assert not ops.move(el(b, "ВВЕДЕНИЕ"), 1).ok
    assert not ops.move(el(b, "1 Анализ"), -1).ok


def test_change_level_with_subtree(b):
    ops = StructureOps(b)
    assert ops.change_level(el(b, "1.1.2 Проблемы"), -1).ok   # пункт → подглава
    assert heads(b)[2:7] == ["1.1 Обзор", "1.1.1 Виды", "1.2 Проблемы", "1.3 Методы", "1.3.1 Подход"]
    assert ops.change_level(el(b, "1.3 Методы"), -1).ok       # подглава → глава (с пунктами)
    assert heads(b)[5:8] == ["2 Методы", "2.1 Подход", "3 Разработка"]
    assert ops.change_level(el(b, "3 Разработка"), 1).ok      # глава → подглава главы 2
    assert heads(b)[5:8] == ["2 Методы", "2.1 Подход", "2.2 Разработка"]
    r = ops.change_level(el(b, "1 Анализ"), 1)                # перед ней нет главы
    assert not r.ok
    r = ops.change_level(el(b, "1 Анализ"), -1)
    assert not r.ok and "глава" in r.message
    assert ops.change_level(el(b, "2 Методы"), 1).ok          # глава с подглавой → подглава с пунктом
    assert heads(b)[5:8] == ["1.3 Методы", "1.3.1 Подход", "1.3.2 Разработка"]
    r = ops.change_level(el(b, "1.3 Методы"), 1)              # пункты стали бы 4-м уровнем
    assert not r.ok and "Глубже" in r.message


def test_lock_blocks_moves(b, tmp_path):
    ops = StructureOps(b)
    target = el(b, "1.1.1 Виды")
    assert ops.set_locked(target, True).ok
    assert not ops.move(target, 1).ok
    assert not ops.change_level(target, -1).ok
    r = ops.move(el(b, "1.1.2 Проблемы"), -1)                # поменяться с закреплённым
    assert not r.ok and "закреплён" in r.message
    assert not ops.delete(target).ok
    path = b.save(tmp_path / "lock.docx")
    b2 = DocumentBuilder.open(path)
    ops2 = StructureOps(b2)
    t2 = next(it.element for it in b2.outline() if it.text == "1.1.1 Виды")
    assert ops2.is_locked(t2)                                  # замок сохранился в файле
    ops2.set_locked(t2, False)
    assert ops2.move(t2, 1).ok


def test_add_element_near_selection(b):
    ops = StructureOps(b)
    r = ops.add_element(Kind.SECTION, near=el(b, "1.1.1 Виды"))
    assert r.ok and "1.1.2 Новый пункт" in heads(b)
    assert heads(b).index("1.1.2 Новый пункт") == heads(b).index("1.1.1 Виды") + 1
    r = ops.add_element(Kind.CHAPTER, near=el(b, "1.2.1 Подход"))
    assert heads(b).index("2 Новая глава") < heads(b).index("3 Разработка")
    r = ops.add_element(Kind.SUBCHAPTER, near=el(b, "2 Новая глава"))
    assert "2.1 Новая подглава" in heads(b)
    avail = {k: ok for k, _, ok, _ in ops.available_elements()}
    assert avail[Kind.INTRODUCTION] is False and avail[Kind.TOC] is True
    assert ops.add_element(Kind.TOC).ok and ops.add_element(Kind.BIBLIOGRAPHY).ok
    assert ops.add_element(Kind.TITLE).ok
    kinds = [it.kind for it in b.outline() if it.level == 0]
    assert kinds[:3] == [Kind.TITLE, Kind.TOC, Kind.INTRODUCTION]
    assert kinds[-1] == Kind.BIBLIOGRAPHY


def test_appendices_as_list_with_references(b, tmp_path):
    b.insert_block(B.APPENDIX, "Таблицы.", heading="Исходные данные")
    b.insert_block(B.APPENDIX, "Код.", heading="Листинг")
    apps = b.appendices.entries()
    assert [(e.key, e.letter, e.title) for e in apps] == [("P1", "А", "Исходные данные"),
                                                          ("P2", "Б", "Листинг")]
    b.insert_block(B.BODY, "Код приведён в приложении {@P2}, данные — в {@P1}.",
                   position=InsertPosition("section_end", el(b, "1 Анализ")))
    texts = [p.text for p in b.document.paragraphs]
    assert "Код приведён в приложении Б, данные — в А." in texts
    assert b.appendices.sort_by_mention() == 2             # ГОСТ: в порядке ссылок
    assert [e.title for e in b.appendices.entries()] == ["Листинг", "Исходные данные"]
    texts = [p.text for p in b.document.paragraphs]
    assert "Код приведён в приложении А, данные — в Б." in texts
    assert "ПРИЛОЖЕНИЕ А" in texts and texts.index("Листинг") < texts.index("Исходные данные")
    StructureOps(b).move(el(b, "ПРИЛОЖЕНИЕ Б"), -1)
    assert "Код приведён в приложении Б, данные — в А." in [p.text for p in b.document.paragraphs]
    c = StructureOps(b).read(el(b, "1.2.1 Подход"))     # текст встал в конец раздела главы 1
    assert "{@P2}" in c.body and "{@P1}" in c.body       # токены восстанавливаются для правки
    path = b.save(tmp_path / "apps.docx")
    assert "ПРИЛОЖЕНИЕ А" in [p.text for p in Document(str(path)).paragraphs]
