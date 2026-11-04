"""Стили Word по профилю оформления.

Принцип: всё оформление живёт в стилях, а вставляемые абзацы получают только
ссылку на стиль (w:pStyle) и «чистые» runs без прямого форматирования. Тогда
документ не «разваливается» при повторном открытии/правке в Word: пользователь
меняет стиль — меняется весь документ, а оглавление и навигация работают.

Заголовки глав/подглав/пунктов — встроенные Heading 1–3 (их понимают
оглавление Word, навигация, LibreOffice, Google Docs). Остальные роли —
собственные стили с латинскими styleId и русскими именами.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import ooxml
from .ooxml import PPR_SEQ, RPR_SEQ
from .profile import FormattingProfile, ParagraphSpec

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


class Role(str, Enum):
    BODY = "body"
    TITLE = "title_page"
    STRUCTURAL = "structural_heading"
    CHAPTER = "chapter_heading"
    SUBCHAPTER = "subchapter_heading"
    SECTION = "section_heading"
    BIBLIOGRAPHY = "bibliography"
    APPENDIX = "appendix_heading"
    APPENDIX_TITLE = "appendix_title"
    TOC_HEADING = "toc_heading"
    FIGURE = "figure"
    FIGURE_CAPTION = "figure_caption"


@dataclass(frozen=True)
class StyleDef:
    style_id: str
    name: str                 # для встроенных — каноническое имя python-docx ("Heading 1")
    builtin: bool = False
    outline_level: int | None = None


STYLE_DEFS: dict[Role, StyleDef] = {
    Role.BODY: StyleDef("GostBody", "ГОСТ Основной текст"),
    Role.TITLE: StyleDef("GostTitlePage", "ГОСТ Титульный лист"),
    Role.STRUCTURAL: StyleDef("GostStructuralHeading", "ГОСТ Структурный элемент", outline_level=0),
    Role.CHAPTER: StyleDef("Heading1", "Heading 1", builtin=True, outline_level=0),
    Role.SUBCHAPTER: StyleDef("Heading2", "Heading 2", builtin=True, outline_level=1),
    Role.SECTION: StyleDef("Heading3", "Heading 3", builtin=True, outline_level=2),
    Role.BIBLIOGRAPHY: StyleDef("GostBibliography", "ГОСТ Список литературы"),
    Role.APPENDIX: StyleDef("GostAppendixHeading", "ГОСТ Приложение", outline_level=0),
    Role.APPENDIX_TITLE: StyleDef("GostAppendixTitle", "ГОСТ Заголовок приложения"),
    # Без outlineLvl: заголовок «СОДЕРЖАНИЕ» сам в оглавление не попадает (ГОСТ 7.32).
    Role.TOC_HEADING: StyleDef("GostTocHeading", "ГОСТ Заголовок содержания"),
    Role.FIGURE: StyleDef("GostFigure", "ГОСТ Рисунок"),
    Role.FIGURE_CAPTION: StyleDef("GostFigureCaption", "ГОСТ Подпись рисунка"),
}

# Профильные атрибуты для ролей, у которых нет собственного ParagraphSpec.
SPEC_ATTR = {Role.TOC_HEADING: "structural_heading"}

ROLE_BY_STYLE_ID = {d.style_id: r for r, d in STYLE_DEFS.items()}
# LibreOffice при пересохранении меняет styleId собственных стилей (GostBody → Style5),
# но сохраняет имя — поэтому роль ищем и по имени.
ROLE_BY_STYLE_NAME = {d.name.lower(): r for r, d in STYLE_DEFS.items() if not d.builtin}
HEADING_ROLES = (Role.STRUCTURAL, Role.CHAPTER, Role.SUBCHAPTER, Role.SECTION, Role.APPENDIX,
                 Role.TOC_HEADING)


class StyleManager:
    def __init__(self, document, profile: FormattingProfile):
        self.document = document
        self.profile = profile

    # ----------------------------------------------------------- поиск стилей
    def find_style(self, style_id: str = None, name: str = None):
        for st in self.document.styles:
            if style_id and st.style_id == style_id:
                return st
        if name:
            for st in self.document.styles:
                if st.name and st.name.lower() == name.lower():
                    return st
            try:
                return self.document.styles[name]
            except KeyError:
                return None
        return None

    def style_for(self, role: Role):
        d = STYLE_DEFS[role]
        st = self.find_style(d.style_id, d.name)
        if st is None:
            st = self._create(role)
            self._apply(st, self._spec_for(role), d.outline_level,
                        is_heading=role in HEADING_ROLES)
        return st

    def _create(self, role: Role):
        d = STYLE_DEFS[role]
        st = self.document.styles.add_style(d.name, WD_STYLE_TYPE.PARAGRAPH)
        if not d.builtin:
            st.style_id = d.style_id
        st.base_style = self.document.styles["Normal"]
        st.quick_style = True
        st.hidden = False
        st.unhide_when_used = False
        return st

    def _spec_for(self, role: Role) -> ParagraphSpec:
        return getattr(self.profile, SPEC_ATTR.get(role, role.value))

    def ensure_toc_styles(self) -> list:
        """Стили пунктов оглавления «toc 1..3» (встроенные имена Word) с отточием до номера."""
        from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

        body = self.profile.body
        pg = self.profile.page
        width_cm = (pg.width_mm - pg.margin_left_mm - pg.margin_right_mm) / 10
        out = []
        for level in range(1, 4):
            name = f"toc {level}"
            st = self.find_style(f"TOC{level}", name)
            if st is None:
                st = self.document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                st.style_id = f"TOC{level}"
                st.base_style = self.document.styles["Normal"]
                st.unhide_when_used = True
                spec = ParagraphSpec(font_name=body.font_name, font_size_pt=body.font_size_pt,
                                     alignment="left", line_spacing=body.line_spacing,
                                     first_line_indent_cm=0, left_indent_cm=0.5 * (level - 1))
                self._apply(st, spec, None)
                pf = st.paragraph_format
                pf.right_indent = Cm(0)
                pf.tab_stops.add_tab_stop(Cm(width_cm), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            out.append(st)
        return out

    # -------------------------------------------------------------- применение
    def ensure_styles(self, overwrite: bool = False) -> None:
        """Создать недостающие стили. overwrite=True — привести существующие к профилю."""
        if overwrite:
            self.apply_document_defaults()
        touched: set[Role] = set()
        for role in Role:
            d = STYLE_DEFS[role]
            st = self.find_style(d.style_id, d.name)
            created = st is None
            if created:
                st = self._create(role)
            if created or overwrite:
                self._apply(st, self._spec_for(role), d.outline_level,
                            is_heading=role in HEADING_ROLES)
                touched.add(role)
        self._link_next_styles(touched)
        if overwrite:
            self._apply_aux_styles()

    def apply_document_defaults(self) -> None:
        """docDefaults и Normal: шрифт, кегль, язык; одинарный интервал без отступов.

        Межстрочный 1,5 и абзацный отступ держим в «ГОСТ Основной текст», а не
        в Normal — иначе они расползутся в таблицы, колонтитулы и подписи.
        """
        body = self.profile.body
        rpr, ppr = ooxml.get_doc_defaults(self.document.styles.element)
        ooxml.set_fonts(rpr, body.font_name)
        ooxml.set_size(rpr, body.font_size_pt)
        ooxml.set_lang(rpr, self.profile.language)
        spacing = ooxml.get_or_insert(ppr, "w:spacing", PPR_SEQ)
        for attr in list(spacing.attrib):
            del spacing.attrib[attr]
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")

        normal = self.document.styles["Normal"]
        n_rpr = normal.element.get_or_add_rPr()
        ooxml.set_fonts(n_rpr, body.font_name)
        ooxml.set_size(n_rpr, body.font_size_pt)
        ooxml.set_black_color(n_rpr)
        ooxml.set_lang(n_rpr, self.profile.language)
        pf = normal.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        pf.first_line_indent = None
        n_ppr = normal.element.get_or_add_pPr()
        ooxml.remove_child(n_ppr, "w:contextualSpacing")

    def _apply_aux_styles(self) -> None:
        """Колонтитулы и оглавление — тем же шрифтом (иначе номер страницы будет Calibri)."""
        body = self.profile.body
        for name in ("Header", "Footer", "TOC Heading", "toc 1", "toc 2", "toc 3"):
            st = self.find_style(name=name)
            if st is None:
                if name in ("Header", "Footer"):
                    st = self.document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                    st.base_style = self.document.styles["Normal"]
                else:
                    continue
            rpr = st.element.get_or_add_rPr()
            ooxml.set_fonts(rpr, body.font_name)
            ooxml.set_size(rpr, body.font_size_pt)
            ooxml.set_black_color(rpr)

    def _link_next_styles(self, roles: set[Role]) -> None:
        """После Enter в заголовке Word продолжит «ГОСТ Основной текст».

        Только для стилей, созданных/обновлённых сейчас: чужие стили не трогаем.
        """
        bd = STYLE_DEFS[Role.BODY]
        body = self.find_style(bd.style_id, bd.name)
        for role in roles:
            d = STYLE_DEFS[role]
            st = self.find_style(d.style_id, d.name)
            if st is None:
                continue
            if role in (Role.BODY, Role.BIBLIOGRAPHY, Role.TITLE):
                st.next_paragraph_style = st
            elif body is not None:
                st.next_paragraph_style = body

    def _apply(self, style, spec: ParagraphSpec, outline_level: int | None,
               is_heading: bool = False) -> None:
        font = style.font
        font.size = Pt(spec.font_size_pt)
        font.bold = spec.bold
        font.italic = spec.italic
        font.all_caps = spec.all_caps
        font.underline = False
        rpr = style.element.get_or_add_rPr()
        ooxml.set_fonts(rpr, spec.font_name)
        ooxml.set_size(rpr, spec.font_size_pt)
        ooxml.set_on_off(rpr, "w:bCs", RPR_SEQ, spec.bold)
        ooxml.set_on_off(rpr, "w:iCs", RPR_SEQ, spec.italic)
        ooxml.set_black_color(rpr)
        ooxml.set_lang(rpr, self.profile.language)

        pf = style.paragraph_format
        pf.alignment = ALIGN_MAP.get(spec.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
        pf.left_indent = Cm(spec.left_indent_cm)
        pf.right_indent = Cm(0)
        pf.first_line_indent = Cm(spec.first_line_indent_cm)
        pf.space_before = Pt(spec.space_before_pt)
        pf.space_after = Pt(spec.space_after_pt)
        pf.line_spacing = spec.line_spacing
        pf.page_break_before = spec.page_break_before
        pf.keep_with_next = spec.keep_with_next
        pf.keep_together = True if is_heading else None
        pf.widow_control = True

        ppr = style.element.get_or_add_pPr()
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None:
            for attr in ("w:beforeAutospacing", "w:afterAutospacing", "w:beforeLines", "w:afterLines"):
                spacing.attrib.pop(qn(attr), None)
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            for attr in ("w:firstLineChars", "w:hangingChars", "w:leftChars", "w:startChars"):
                ind.attrib.pop(qn(attr), None)
        ooxml.remove_child(ppr, "w:contextualSpacing")
        if spec.no_hyphenation:
            ooxml.set_on_off(ppr, "w:suppressAutoHyphens", PPR_SEQ, True)
        else:
            ooxml.remove_child(ppr, "w:suppressAutoHyphens")
        if outline_level is not None:
            ooxml.get_or_insert(ppr, "w:outlineLvl", PPR_SEQ).set(qn("w:val"), str(outline_level))
            # Текстовая нумерация + многоуровневый список Word = двойные номера.
            if self.profile.numbering.enabled:
                ooxml.remove_child(ppr, "w:numPr")
