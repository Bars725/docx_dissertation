import pytest
from docx import Document
from docx.oxml.ns import qn

from dissertation.document_builder import BlockType as B, DocumentBuilder, InsertPosition
from dissertation.figures import FigureSpec, Figures, split_figure_tokens
from dissertation.gost import gost_7_32_2017
from dissertation.structure_ops import StructureOps

PIL = pytest.importorskip("PIL.Image")


@pytest.fixture
def images(tmp_path):
    paths = {}
    for name, size in (("scheme.png", (1600, 900)), ("tall.png", (400, 3000)), ("small.png", (200, 100))):
        p = tmp_path / name
        PIL.new("RGB", size, (90, 140, 200)).save(p, dpi=(96, 96))
        paths[name] = str(p)
    return paths


def texts(b):
    return [p.text for p in b.document.paragraphs]


def test_token_parsing():
    segs = split_figure_tokens("До.\n⟦Рисунок: Схема системы. | C:\\img\\a.png | 12,5 см⟧\nПосле.")
    assert segs[0] == "До.\n" and segs[2] == "\nПосле."
    assert segs[1] == FigureSpec("C:\\img\\a.png", "Схема системы", 12.5)
    assert FigureSpec("x.png", "Т", None).token() == "⟦Рисунок: Т | x.png⟧"


def test_insert_figure_with_caption_and_numbering(images, tmp_path):
    b = DocumentBuilder.new(gost_7_32_2017())
    b.insert_block(B.CHAPTER, f"Текст до.\n⟦Рисунок: схема системы. | {images['scheme.png']}⟧",
                   heading="Обзор")
    assert Figures(b).entries()[0].key == "F1"
    b.insert_block(B.BODY, "Как видно на рисунке {@F1}, всё работает.")
    t = texts(b)
    assert "Рисунок 1 – Схема системы" in t                # ГОСТ: с прописной, без точки, через тире
    assert "Как видно на рисунке 1, всё работает." in t
    i = t.index("Рисунок 1 – Схема системы")
    img_p = b.document.paragraphs[i - 1]
    assert img_p.style.name == "ГОСТ Рисунок" and img_p._p.find(".//" + qn("w:drawing")) is not None
    assert img_p.paragraph_format.keep_with_next or img_p.style.paragraph_format.keep_with_next
    # второй рисунок вставлен ВЫШЕ — номера и ссылка пересчитались
    first = next(it.element for it in b.outline())
    b.insert_block(B.BODY, f"⟦Рисунок: Маленький | {images['small.png']}⟧",
                   position=InsertPosition("after", first))
    t = texts(b)
    assert t.index("Рисунок 1 – Маленький") < t.index("Рисунок 2 – Схема системы")
    assert "Как видно на рисунке 2, всё работает." in t
    path = b.save(tmp_path / "fig.docx")
    assert "Рисунок 2 – Схема системы" in [p.text for p in Document(str(path)).paragraphs]


def test_figure_size_fits_page(images):
    b = DocumentBuilder.new(gost_7_32_2017())
    b.insert_block(B.CHAPTER, f"⟦Рисунок: Широкий | {images['scheme.png']}⟧\n"
                              f"⟦Рисунок: Высокий | {images['tall.png']}⟧\n"
                              f"⟦Рисунок: Ширина | {images['small.png']} | 8 см⟧", heading="Г")
    shapes = b.document.inline_shapes
    pg = b.profile.page
    text_w = (pg.width_mm - pg.margin_left_mm - pg.margin_right_mm) / 10
    assert shapes[0].width.cm <= text_w + 0.01
    assert shapes[1].height.cm <= (pg.height_mm - pg.margin_top_mm - pg.margin_bottom_mm) / 10 * 0.66
    assert abs(shapes[2].width.cm - 8) < 0.05


def test_figures_in_appendix_and_chapter_mode(images):
    p = gost_7_32_2017()
    p.figure_numbering = "chapter"
    b = DocumentBuilder.new(p)
    b.insert_block(B.CHAPTER, f"⟦Рисунок: Первый | {images['small.png']}⟧", heading="Один")
    b.insert_block(B.CHAPTER, f"⟦Рисунок: Второй | {images['small.png']}⟧\n"
                              f"⟦Рисунок: Третий | {images['small.png']}⟧", heading="Два")
    b.insert_block(B.APPENDIX, f"⟦Рисунок: В приложении | {images['small.png']}⟧", heading="Данные")
    t = texts(b)
    assert {"Рисунок 1.1 – Первый", "Рисунок 2.1 – Второй", "Рисунок 2.2 – Третий",
            "Рисунок А.1 – В приложении"} <= set(t)


def test_figure_is_marker_when_editing(images):
    b = DocumentBuilder.new(gost_7_32_2017())
    b.insert_block(B.CHAPTER, f"Текст.\n⟦Рисунок: Схема | {images['small.png']}⟧\nЕщё текст.",
                   heading="Глава")
    head = b.outline()[0].element
    ops = StructureOps(b)
    c = ops.read(head)
    assert c.body == "Текст.\n\n⟦#1 Рисунок 1 – Схема⟧\n\nЕщё текст."
    # переименовать подпись прямо в маркере и переставить рисунок в конец
    ops.update(head, None, "Текст.\n\nЕщё текст.\n\n⟦#1 Рисунок 1 – Структурная схема⟧")
    t = texts(b)
    assert t[-1] == "Рисунок 1 – Структурная схема"
    assert len(b.document.inline_shapes) == 1
    assert Figures(b).entries()[0].key == "F1"                 # ключ для ссылок сохранился


def test_missing_image_file_reported():
    b = DocumentBuilder.new(gost_7_32_2017())
    with pytest.raises(ValueError, match="не найден"):
        b.insert_block(B.CHAPTER, "⟦Рисунок: Нет | /нет/такого/файла.png⟧", heading="Г")
    pr = b.preview_block(B.CHAPTER, "⟦Рисунок: Нет | /нет/такого/файла.png⟧", heading="Г")
    assert "ФАЙЛ НЕ НАЙДЕН" in pr.items[1][1]
    assert pr.items[2][1] == "Рисунок 1 – Нет"
