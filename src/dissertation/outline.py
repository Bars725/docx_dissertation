"""Распознавание структуры документа: какие абзацы — главы, приложения, список литературы.

Работает и для «своих» документов (по styleId наших стилей), и для чужих
(по уровню структуры outlineLvl в цепочке стилей + тексту заголовка).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import ooxml
from .runs import visible_text
from .style_manager import ROLE_BY_STYLE_ID, ROLE_BY_STYLE_NAME, Role


class Kind(str, Enum):
    TITLE = "title"
    TOC = "toc"
    INTRODUCTION = "introduction"
    CHAPTER = "chapter"
    SUBCHAPTER = "subchapter"
    SECTION = "section"
    CONCLUSION = "conclusion"
    BIBLIOGRAPHY = "bibliography"
    APPENDIX = "appendix"
    OTHER_STRUCTURAL = "other_structural"


STRUCTURAL_PATTERNS = [
    (Kind.TOC, re.compile(r"^(СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ)$", re.I)),
    (Kind.INTRODUCTION, re.compile(r"^ВВЕДЕНИЕ$", re.I)),
    (Kind.CONCLUSION, re.compile(r"^ЗАКЛЮЧЕНИЕ$", re.I)),
    (Kind.BIBLIOGRAPHY, re.compile(r"^(СПИСОК\s+(ИСПОЛЬЗОВАННЫХ\s+)?(ЛИТЕРАТУРЫ|ИСТОЧНИКОВ)"
                                   r"|БИБЛИОГРАФИЧЕСКИЙ\s+СПИСОК|БИБЛИОГРАФИЯ)", re.I)),
    (Kind.APPENDIX, re.compile(r"^ПРИЛОЖЕНИ[ЕЯ]\b", re.I)),
    (Kind.OTHER_STRUCTURAL, re.compile(
        r"^(РЕФЕРАТ|АННОТАЦИЯ|ОПРЕДЕЛЕНИЯ|ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ|ПЕРЕЧЕНЬ СОКРАЩЕНИЙ.*"
        r"|СПИСОК СОКРАЩЕНИЙ.*|СЛОВАРЬ ТЕРМИНОВ|СПИСОК ИСПОЛНИТЕЛЕЙ|ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ"
        r"|СПИСОК ИЛЛЮСТРАТИВНОГО МАТЕРИАЛА)$", re.I)),
]

_HEADING_NAME = re.compile(r"^heading\s*([1-9])$", re.I)


def structural_kind(text: str) -> Kind | None:
    t = " ".join(text.split())
    for kind, rx in STRUCTURAL_PATTERNS:
        if rx.match(t):
            return kind
    return None


@dataclass
class OutlineItem:
    element: object          # w:p
    kind: Kind
    level: int               # 0..2
    text: str
    style_id: str | None

    @property
    def label(self) -> str:
        return ("    " * self.level) + (self.text or "(пусто)")


class StyleResolver:
    """Эффективные свойства абзаца по цепочке: прямые → стиль → basedOn → docDefaults."""

    def __init__(self, document):
        self.styles_el = document.styles.element
        self._by_id = {s.get(qn("w:styleId")): s for s in self.styles_el.findall(qn("w:style"))}
        self._default_para = None
        self._roles: dict[str, Role] = {}
        for sid, st in self._by_id.items():
            name_el = st.find(qn("w:name"))
            name = (name_el.get(qn("w:val")) if name_el is not None else "").lower()
            role = ROLE_BY_STYLE_ID.get(sid) or ROLE_BY_STYLE_NAME.get(name)
            if role is not None and (role not in self._roles.values() or sid in ROLE_BY_STYLE_ID):
                self._roles[sid] = role
        for s in self.styles_el.findall(qn("w:style")):
            if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true"):
                self._default_para = s.get(qn("w:styleId"))

    def style_id_of(self, p_el) -> str | None:
        ppr = p_el.find(qn("w:pPr"))
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                return ps.get(qn("w:val"))
        return self._default_para

    def role_of(self, p_el) -> Role | None:
        """Роль абзаца по нашему стилю (styleId или имя), иначе None."""
        return self._roles.get(self.style_id_of(p_el) or "")

    def style_chain(self, style_id: str | None) -> list:
        chain, seen = [], set()
        while style_id and style_id not in seen and style_id in self._by_id:
            seen.add(style_id)
            st = self._by_id[style_id]
            chain.append(st)
            based = st.find(qn("w:basedOn"))
            style_id = based.get(qn("w:val")) if based is not None else None
        return chain

    def style_name(self, style_id: str | None) -> str:
        st = self._by_id.get(style_id or "")
        if st is None:
            return ""
        name = st.find(qn("w:name"))
        return name.get(qn("w:val")) if name is not None else ""

    def ppr_chain(self, p_el) -> list:
        """Список w:pPr от самого приоритетного к наименее приоритетному."""
        out = []
        ppr = p_el.find(qn("w:pPr"))
        if ppr is not None:
            out.append(ppr)
        for st in self.style_chain(self.style_id_of(p_el)):
            sp = st.find(qn("w:pPr"))
            if sp is not None:
                out.append(sp)
        dd = self.styles_el.find(qn("w:docDefaults"))
        if dd is not None:
            pd = dd.find(qn("w:pPrDefault"))
            d = pd.find(qn("w:pPr")) if pd is not None else None
            if d is not None:
                out.append(d)
        return out

    def outline_level(self, p_el) -> int | None:
        for ppr in self.ppr_chain(p_el):
            ol = ppr.find(qn("w:outlineLvl"))
            if ol is not None:
                v = int(ol.get(qn("w:val"), "9"))
                return v if v < 9 else None
        # Стили «heading N» без outlineLvl (бывает в файлах из Google Docs/LO).
        for st in self.style_chain(self.style_id_of(p_el)):
            name = st.find(qn("w:name"))
            m = _HEADING_NAME.match(name.get(qn("w:val")) if name is not None else "")
            if m:
                return int(m.group(1)) - 1
        return None


def classify_paragraph(p_el, resolver: StyleResolver) -> OutlineItem | None:
    style_id = resolver.style_id_of(p_el)
    text = " ".join(ooxml.element_text(p_el).split())
    role = resolver.role_of(p_el)
    if role == Role.TITLE:
        return OutlineItem(p_el, Kind.TITLE, 0, text, style_id)
    if role == Role.APPENDIX:
        return OutlineItem(p_el, Kind.APPENDIX, 0, text, style_id)
    if role == Role.TOC_HEADING:
        return OutlineItem(p_el, Kind.TOC, 0, text, style_id)
    if role == Role.STRUCTURAL:
        return OutlineItem(p_el, structural_kind(text) or Kind.OTHER_STRUCTURAL, 0, text, style_id)
    level = resolver.outline_level(p_el)
    if level is None or level > 2 or not text:
        return None
    if level == 0:
        kind = structural_kind(text)
        return OutlineItem(p_el, kind or Kind.CHAPTER, 0, text, style_id)
    return OutlineItem(p_el, Kind.SUBCHAPTER if level == 1 else Kind.SECTION, level, text, style_id)


TITLE_BOOKMARK_PREFIX = "_GostTitle_"


def is_title_table(tbl_el) -> bool:
    """Таблица-разметка титульного листа, созданная формой (по закладкам полей)."""
    for b in tbl_el.iter(qn("w:bookmarkStart")):
        if (b.get(qn("w:name")) or "").startswith(TITLE_BOOKMARK_PREFIX):
            return True
    return False


def build_outline(document) -> list[OutlineItem]:
    resolver = StyleResolver(document)
    items = []
    for el in ooxml.iter_body_blocks(document.element.body):
        if el.tag == ooxml.W_TBL:
            if is_title_table(el):
                items.append(OutlineItem(el, Kind.TITLE, 0, "[Титульный лист]", None))
            continue
        if el.tag != ooxml.W_P:
            continue
        item = classify_paragraph(el, resolver)
        if item is not None:
            items.append(item)
    return items


def paragraph(el, document) -> Paragraph:
    return Paragraph(el, document._body)


def text_of(el, document) -> str:
    return visible_text(paragraph(el, document))
