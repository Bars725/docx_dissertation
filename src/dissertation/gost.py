"""Встроенные пресеты ГОСТ.

ГОСТ 7.32-2017 (отчёт о НИР) + ГОСТ 2.105-95 — пресет по умолчанию.
ГОСТ Р 7.0.11-2011 («Диссертация и автореферат диссертации») — альтернативный
пресет: именно его обычно требуют диссертационные советы.

Значения, которые стандарт не фиксирует жёстко (кегль 12–14, интервалы вокруг
заголовков), выбраны по распространённой практике и вынесены в профиль —
их легко поменять под требования конкретного вуза.
"""
from __future__ import annotations

from .profile import (
    FormattingProfile,
    HeadingNumbering,
    PageNumbering,
    PageSetup,
    ParagraphSpec,
)

FONT = "Times New Roman"
BODY_PT = 14.0


def _body(**kw) -> ParagraphSpec:
    base = dict(
        font_name=FONT, font_size_pt=BODY_PT, alignment="justify",
        line_spacing=1.5, first_line_indent_cm=1.25,
    )
    base.update(kw)
    return ParagraphSpec(**base)


def gost_7_32_2017() -> FormattingProfile:
    """ГОСТ 7.32-2017.

    Поля: левое 30, правое 15, верхнее/нижнее 20 мм. TNR, ≥12 пт (берём 14),
    интервал 1,5, абзацный отступ 1,25 см, по ширине. Номер страницы — внизу
    по центру, без точки; на титульном не ставится. Структурные элементы —
    прописными, по центру, полужирным, с новой страницы. Заголовки разделов —
    с абзацного отступа, полужирным, без точки; раздел — с новой страницы.
    """
    heading = dict(bold=True, line_spacing=1.5, keep_with_next=True, no_hyphenation=True)
    return FormattingProfile(
        name="ГОСТ 7.32-2017",
        source="builtin:gost-7.32-2017",
        page=PageSetup(margin_left_mm=30, margin_right_mm=15, margin_top_mm=20, margin_bottom_mm=20),
        body=_body(),
        title_page=_body(alignment="center", first_line_indent_cm=0, line_spacing=1.0),
        structural_heading=_body(alignment="center", first_line_indent_cm=0, all_caps=False,
                                 page_break_before=True, space_after_pt=12, **heading),
        chapter_heading=_body(alignment="justify", first_line_indent_cm=1.25,
                              page_break_before=True, space_after_pt=12, **heading),
        subchapter_heading=_body(alignment="justify", first_line_indent_cm=1.25,
                                 space_before_pt=12, space_after_pt=12, **heading),
        section_heading=_body(alignment="justify", first_line_indent_cm=1.25,
                              space_before_pt=12, space_after_pt=6, **heading),
        bibliography=_body(),
        appendix_heading=_body(alignment="center", first_line_indent_cm=0,
                               page_break_before=True, **heading),
        appendix_title=_body(alignment="center", first_line_indent_cm=0, space_after_pt=12,
                             **heading),
        numbering=HeadingNumbering(enabled=True, chapter="{n} {title}",
                                   subchapter="{n}.{m} {title}", section="{n}.{m}.{k} {title}"),
        page_numbering=PageNumbering(enabled=True, position="footer", alignment="center",
                                     hide_on_first_page=True),
        structural_uppercase=True,
        bibliography_title="СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
        appendix_label="ПРИЛОЖЕНИЕ {letter}",
    )


def gost_r_7_0_11_2011() -> FormattingProfile:
    """ГОСТ Р 7.0.11-2011 (диссертация).

    Поля: левое 25, правое 10, верхнее/нижнее 20 мм. 12–14 пт, интервал 1,5,
    абзацный отступ «5 знаков» (≈1,25 см). Номер страницы — на середине
    верхнего поля, титульный лист не нумеруется. Главы — с новой страницы,
    заголовки посередине, без точки, без переносов, отделены от текста
    «тремя интервалами» (здесь 18 пт). Приложения — «Приложение А» вверху
    посередине, заголовок отдельной строкой. Список литературы — ГОСТ Р 7.0.100-2018.
    """
    heading = dict(bold=True, line_spacing=1.5, keep_with_next=True, no_hyphenation=True,
                   alignment="center", first_line_indent_cm=0)
    return FormattingProfile(
        name="ГОСТ Р 7.0.11-2011 (диссертация)",
        source="builtin:gost-r-7.0.11-2011",
        page=PageSetup(margin_left_mm=25, margin_right_mm=10, margin_top_mm=20, margin_bottom_mm=20),
        body=_body(),
        title_page=_body(alignment="center", first_line_indent_cm=0, line_spacing=1.0),
        structural_heading=_body(page_break_before=True, space_after_pt=18, **heading),
        chapter_heading=_body(page_break_before=True, space_after_pt=18, **heading),
        subchapter_heading=_body(space_before_pt=18, space_after_pt=18, **heading),
        section_heading=_body(space_before_pt=12, space_after_pt=12, **heading),
        bibliography=_body(),
        appendix_heading=_body(page_break_before=True, **heading),
        appendix_title=_body(space_after_pt=18, **heading),
        numbering=HeadingNumbering(enabled=True, chapter="Глава {n}. {title}",
                                   subchapter="{n}.{m} {title}", section="{n}.{m}.{k} {title}"),
        page_numbering=PageNumbering(enabled=True, position="header", alignment="center",
                                     hide_on_first_page=True),
        structural_uppercase=True,
        bibliography_title="СПИСОК ЛИТЕРАТУРЫ",
        appendix_label="Приложение {letter}",
    )


PRESETS = {
    "gost-7.32-2017": gost_7_32_2017,
    "gost-r-7.0.11-2011": gost_r_7_0_11_2011,
}

DEFAULT_PRESET = "gost-7.32-2017"


def default_profile() -> FormattingProfile:
    return PRESETS[DEFAULT_PRESET]()
