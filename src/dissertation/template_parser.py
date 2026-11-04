"""Извлечение профиля оформления из образца .docx.

Ключевая сложность: python-docx отдаёт только *локальные* значения
(``paragraph.paragraph_format.first_line_indent`` = None, если отступ задан в
стиле или в basedOn-родителе). Поэтому параметры вычисляются вручную по
цепочке приоритетов OOXML:

    прямое форматирование run → символьный стиль → стиль абзаца → basedOn … →
    docDefaults → шрифты темы (theme1.xml)

Значения собираются по многим абзацам и выбираются взвешенной модой
(вес — длина текста), чтобы одиночные «криво» оформленные абзацы не
перекосили результат. Всё, что найти не удалось, остаётся из пресета ГОСТ,
и это фиксируется в ``profile.extraction_report``.
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import replace
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from lxml import etree

from . import ooxml
from .gost import default_profile
from .outline import Kind, StyleResolver, classify_paragraph
from .profile import FormattingProfile, ParagraphSpec

TWIPS_PER_CM = 566.929
JC_MAP = {"both": "justify", "distribute": "justify", "center": "center", "right": "right",
          "end": "right", "left": "left", "start": "left"}
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NON_BODY_STYLE = re.compile(r"(toc|caption|title|header|footer|list|footnote|table|quote|"
                            r"оглавл|назван|подпис|колонтитул|список|таблиц)", re.I)


def _weighted_mode(pairs):
    """[(value, weight)] → значение с наибольшим суммарным весом (None игнорируется)."""
    c = Counter()
    for value, weight in pairs:
        if value is not None:
            c[value] += max(weight, 1)
    return c.most_common(1)[0][0] if c else None


class TemplateParser:
    def __init__(self, document, base: FormattingProfile | None = None, source: str = ""):
        self.document = document
        self.base = base or default_profile()
        self.source = source
        self.resolver = StyleResolver(document)
        self.report: list[str] = []
        self._theme_fonts = self._read_theme_fonts()

    @classmethod
    def from_path(cls, path: str | Path, base: FormattingProfile | None = None) -> "TemplateParser":
        return cls(Document(str(path)), base=base, source=str(path))

    # ================================================================ главное
    def extract(self) -> FormattingProfile:
        prof = self.base.copy()
        prof.source = self.source or "document"
        prof.name = f"Образец: {Path(self.source).stem}" if self.source else "Образец"
        self._extract_page(prof)

        body_paras, groups = self._collect()
        prof.body = self._spec(body_paras, prof.body, "Основной текст")
        prof.structural_heading = self._spec(groups["structural"], prof.structural_heading,
                                             "Заголовок структурного элемента", heading=True)
        prof.chapter_heading = self._spec(groups[Kind.CHAPTER], prof.chapter_heading,
                                          "Заголовок главы", heading=True)
        prof.subchapter_heading = self._spec(groups[Kind.SUBCHAPTER], prof.subchapter_heading,
                                             "Заголовок подглавы", heading=True)
        prof.section_heading = self._spec(groups[Kind.SECTION], prof.section_heading,
                                          "Заголовок пункта", heading=True)
        prof.appendix_heading = self._spec(groups[Kind.APPENDIX], prof.appendix_heading,
                                           "Обозначение приложения", heading=True)
        prof.bibliography = self._spec(groups["bib_entries"], prof.bibliography,
                                       "Список литературы")
        prof.title_page = self._spec(groups["title"], prof.title_page, "Титульный лист")
        # Заголовок приложения и титул — наследуют шрифт основного текста, если не найдены.
        if not groups[Kind.APPENDIX]:
            prof.appendix_title = replace(prof.appendix_title, font_name=prof.body.font_name,
                                          font_size_pt=prof.body.font_size_pt)

        self._extract_texts(prof, groups)
        self._extract_numbering(prof, groups)
        self._extract_page_numbering(prof)
        prof.extraction_report = self.report
        return prof

    # ============================================================== страница
    def _extract_page(self, prof: FormattingProfile) -> None:
        sections = self.document.sections
        if not sections:
            self.report.append("Параметры страницы: секций нет — по ГОСТ")
            return
        s = sections[0]
        pg = prof.page
        for attr, length in (("width_mm", s.page_width), ("height_mm", s.page_height),
                             ("margin_left_mm", s.left_margin), ("margin_right_mm", s.right_margin),
                             ("margin_top_mm", s.top_margin), ("margin_bottom_mm", s.bottom_margin),
                             ("header_distance_mm", s.header_distance),
                             ("footer_distance_mm", s.footer_distance)):
            if length is not None:
                setattr(pg, attr, round(length.mm, 1))
        self.report.append(
            f"Страница: {pg.width_mm:g}×{pg.height_mm:g} мм, поля л/п/в/н "
            f"{pg.margin_left_mm:g}/{pg.margin_right_mm:g}/{pg.margin_top_mm:g}/{pg.margin_bottom_mm:g} мм"
            + (f" (секций: {len(sections)}, взята первая)" if len(sections) > 1 else ""))

    # ======================================================== сбор абзацев
    def _collect(self):
        body_paras = []
        groups = defaultdict(list)
        blocks = list(ooxml.iter_body_blocks(self.document.element.body))
        seen_heading = False
        in_bib = False
        prev = None
        for el in blocks:
            if el.tag != ooxml.W_P:
                prev = el
                continue
            item = classify_paragraph(el, self.resolver)
            text = ooxml.element_text(el).strip()
            if item is not None and item.kind != Kind.TITLE:
                seen_heading = True
                in_bib = item.kind == Kind.BIBLIOGRAPHY
                if item.kind in (Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION, Kind.APPENDIX):
                    groups[item.kind].append((el, prev))
                else:
                    groups["structural"].append((el, prev))
                groups["items"].append(item)
            elif text:
                if not seen_heading and len(text) < 200:
                    groups["title"].append((el, prev))
                elif in_bib:
                    groups["bib_entries"].append((el, prev))
                elif self._is_body_candidate(el, text):
                    body_paras.append((el, prev))
            prev = el
        # Без заголовков в образце «титул» = весь документ — не путаем с титулом.
        if not seen_heading:
            groups["title"] = []
        return body_paras, groups

    def _is_body_candidate(self, el, text: str) -> bool:
        if len(text) < 60:
            return False
        sid = self.resolver.style_id_of(el)
        name = self.resolver.style_name(sid)
        if NON_BODY_STYLE.search(sid or "") or NON_BODY_STYLE.search(name):
            return False
        ppr = el.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:numPr")) is not None:
            return False
        return True

    # ===================================================== вычисление свойств
    def _ppr_attr(self, el, tag: str, attr: str):
        for ppr in self.resolver.ppr_chain(el):
            child = ppr.find(qn(tag))
            if child is not None and child.get(qn(attr)) is not None:
                return child.get(qn(attr))
        return None

    def _ppr_onoff(self, el, tag: str):
        for ppr in self.resolver.ppr_chain(el):
            v = ooxml.read_on_off(ppr.find(qn(tag)))
            if v is not None:
                return v
        return False

    def _rpr_chain(self, run_el, p_el) -> list:
        out = []
        rpr = run_el.find(qn("w:rPr"))
        if rpr is not None:
            out.append(rpr)
            rs = rpr.find(qn("w:rStyle"))
            if rs is not None:
                for st in self.resolver.style_chain(rs.get(qn("w:val"))):
                    r = st.find(qn("w:rPr"))
                    if r is not None:
                        out.append(r)
        for st in self.resolver.style_chain(self.resolver.style_id_of(p_el)):
            r = st.find(qn("w:rPr"))
            if r is not None:
                out.append(r)
        dd = self.document.styles.element.find(qn("w:docDefaults"))
        if dd is not None:
            rd = dd.find(qn("w:rPrDefault"))
            r = rd.find(qn("w:rPr")) if rd is not None else None
            if r is not None:
                out.append(r)
        return out

    def _font_of(self, chain) -> str | None:
        # Кириллица попадает в слот hAnsi, латиница — в ascii.
        for slot, theme_slot in (("w:hAnsi", "w:hAnsiTheme"), ("w:ascii", "w:asciiTheme")):
            for rpr in chain:
                rf = rpr.find(qn("w:rFonts"))
                if rf is None:
                    continue
                theme = rf.get(qn(theme_slot))
                if theme:
                    return self._theme_fonts.get("major" if theme.startswith("major") else "minor")
                if rf.get(qn(slot)):
                    return rf.get(qn(slot))
        return None

    @staticmethod
    def _first_val(chain, tag):
        for rpr in chain:
            el = rpr.find(qn(tag))
            if el is not None:
                return el
        return None

    def _run_props(self, p_el):
        """Взвешенная мода свойств символов по runs абзаца."""
        fonts, sizes, bolds, italics, caps = [], [], [], [], []
        for r in p_el.iter(qn("w:r")):
            text = "".join(t.text or "" for t in r.findall(qn("w:t")))
            if not text.strip():
                continue
            w = len(text)
            chain = self._rpr_chain(r, p_el)
            fonts.append((self._font_of(chain), w))
            sz = self._first_val(chain, "w:sz")
            sizes.append((float(sz.get(qn("w:val"))) / 2 if sz is not None else 10.0, w))
            for tag, acc in (("w:b", bolds), ("w:i", italics), ("w:caps", caps)):
                acc.append((bool(ooxml.read_on_off(self._first_val(chain, tag))), w))
        return (_weighted_mode(fonts), _weighted_mode(sizes), _weighted_mode(bolds),
                _weighted_mode(italics), _weighted_mode(caps))

    def _para_props(self, el, prev, font_size: float | None) -> dict:
        jc = self._ppr_attr(el, "w:jc", "w:val")
        first = self._ppr_attr(el, "w:ind", "w:firstLine")
        hanging = self._ppr_attr(el, "w:ind", "w:hanging")
        first_chars = self._ppr_attr(el, "w:ind", "w:firstLineChars")
        left = self._ppr_attr(el, "w:ind", "w:left") or self._ppr_attr(el, "w:ind", "w:start")
        if first is not None:
            first_cm = int(first) / TWIPS_PER_CM
        elif hanging is not None:
            first_cm = -int(hanging) / TWIPS_PER_CM
        elif first_chars is not None and font_size:
            first_cm = int(first_chars) / 100 * font_size / 72 * 2.54
        else:
            first_cm = 0.0
        line = self._ppr_attr(el, "w:spacing", "w:line")
        rule = self._ppr_attr(el, "w:spacing", "w:lineRule") or "auto"
        if line is None:
            spacing = 1.0
        elif rule == "auto":
            spacing = int(line) / 240
        else:  # exact/atLeast: переводим в множитель относительно кегля
            spacing = (int(line) / 20) / ((font_size or 12) * 1.15)
        before = self._ppr_attr(el, "w:spacing", "w:before")
        after = self._ppr_attr(el, "w:spacing", "w:after")
        page_break = self._ppr_onoff(el, "w:pageBreakBefore")
        if not page_break and prev is not None:
            page_break = ooxml.contains_page_break(prev) or (
                prev.tag == ooxml.W_P and ooxml.paragraph_has_section_break(prev))
        return dict(
            alignment=JC_MAP.get(jc or "left", "left"),
            first_line_indent_cm=round(first_cm, 2),
            left_indent_cm=round(int(left) / TWIPS_PER_CM, 2) if left else 0.0,
            line_spacing=round(spacing * 4) / 4,     # к ближайшим 0,25
            space_before_pt=round(int(before) / 20, 1) if before else 0.0,
            space_after_pt=round(int(after) / 20, 1) if after else 0.0,
            page_break_before=page_break,
            keep_with_next=self._ppr_onoff(el, "w:keepNext"),
        )

    def _spec(self, paras, base: ParagraphSpec, label: str, heading: bool = False) -> ParagraphSpec:
        if not paras:
            self.report.append(f"{label}: в образце не найден — по ГОСТ")
            return base
        values = defaultdict(list)
        for el, prev in paras:
            weight = len(ooxml.element_text(el)) or 1
            font, size, bold, italic, caps = self._run_props(el)
            for k, v in (("font_name", font), ("font_size_pt", size), ("bold", bold),
                         ("italic", italic), ("all_caps", caps)):
                values[k].append((v, weight))
            for k, v in self._para_props(el, prev, size).items():
                values[k].append((v, 1 if heading else weight))
        found = {k: _weighted_mode(v) for k, v in values.items()}
        found = {k: v for k, v in found.items() if v is not None}
        spec = replace(base, **found)
        if heading:
            spec.no_hyphenation = True
            spec.keep_with_next = True
        self.report.append(
            f"{label}: {spec.font_name} {spec.font_size_pt:g} пт"
            f"{', полужирный' if spec.bold else ''}, {_ALIGN_RU[spec.alignment]}, "
            f"интервал {spec.line_spacing:g}, отступ {spec.first_line_indent_cm:g} см"
            f"{', с новой страницы' if spec.page_break_before else ''} (по {len(paras)} абз.)")
        return spec

    # ============================================================ тексты/нумерация
    def _extract_texts(self, prof: FormattingProfile, groups) -> None:
        structural = [it for it in groups["items"] if it.kind in (
            Kind.INTRODUCTION, Kind.CONCLUSION, Kind.BIBLIOGRAPHY, Kind.TOC, Kind.OTHER_STRUCTURAL)]
        if structural:
            prof.structural_uppercase = all(it.text == it.text.upper() for it in structural)
        for it in groups["items"]:
            if it.kind == Kind.BIBLIOGRAPHY:
                prof.bibliography_title = it.text
            elif it.kind == Kind.INTRODUCTION:
                prof.introduction_title = it.text
            elif it.kind == Kind.CONCLUSION:
                prof.conclusion_title = it.text
        appendices = [it for it in groups["items"] if it.kind == Kind.APPENDIX]
        if appendices:
            word = appendices[0].text.split()[0]
            prof.appendix_label = f"{word} {{letter}}"
        entries = [ooxml.element_text(el).strip() for el, _ in groups["bib_entries"]]
        if entries:
            numbered_by_list = sum(
                1 for el, _ in groups["bib_entries"]
                if el.find(qn("w:pPr")) is not None and el.find(qn("w:pPr")).find(qn("w:numPr")) is not None)
            fmt = _weighted_mode([(_bib_format(e), 1) for e in entries])
            if fmt:
                prof.bibliography_number_format = fmt
            self.report.append(
                f"Список литературы: {len(entries)} записей, нумерация "
                + ("списком Word (новые записи получат текстовые номера)" if numbered_by_list
                   else f"«{prof.bibliography_number_format.strip()}»"))

    def _extract_numbering(self, prof: FormattingProfile, groups) -> None:
        chapters = [ooxml.element_text(el).strip() for el, _ in groups[Kind.CHAPTER]]
        subs = [ooxml.element_text(el).strip() for el, _ in groups[Kind.SUBCHAPTER]]
        secs = [ooxml.element_text(el).strip() for el, _ in groups[Kind.SECTION]]
        if not chapters:
            self.report.append("Нумерация заголовков: глав нет — формат по ГОСТ")
            return
        word_numbering = any(self._has_numpr(el) for el, _ in groups[Kind.CHAPTER])
        fmt = _weighted_mode([(_heading_format(t, 1), 1) for t in chapters])
        if fmt is None:
            prof.numbering.enabled = False
            self.report.append(
                "Нумерация заголовков: " + ("многоуровневый список Word — номера не добавляются"
                                            if word_numbering else "в образце без номеров — номера не добавляются"))
            return
        prof.numbering.enabled = True
        prof.numbering.chapter = fmt
        sub_fmt = _weighted_mode([(_heading_format(t, 2), 1) for t in subs])
        sec_fmt = _weighted_mode([(_heading_format(t, 3), 1) for t in secs])
        if sub_fmt:
            prof.numbering.subchapter = sub_fmt
        if sec_fmt:
            prof.numbering.section = sec_fmt
        self.report.append(f"Нумерация заголовков: «{prof.numbering.chapter}», "
                           f"«{prof.numbering.subchapter}»")

    def _has_numpr(self, el) -> bool:
        return any(ppr.find(qn("w:numPr")) is not None for ppr in self.resolver.ppr_chain(el))

    # ======================================================= номера страниц
    def _extract_page_numbering(self, prof: FormattingProfile) -> None:
        if not self.document.sections:
            return
        s = self.document.sections[0]
        pn = prof.page_numbering
        found = None
        for position, hf in (("footer", s.footer), ("header", s.header)):
            if not hf.is_linked_to_previous and ooxml.has_page_field(hf._element):
                found = (position, hf)
                break
        first_page = s.different_first_page_header_footer
        if found is None and first_page:
            for position, hf in (("footer", s.first_page_footer), ("header", s.first_page_header)):
                if not hf.is_linked_to_previous and ooxml.has_page_field(hf._element):
                    found = (position, hf)
                    break
        if found is None:
            pn.enabled = False
            self.report.append("Нумерация страниц: в образце не найдена — отключена")
            return
        position, hf = found
        pn.enabled = True
        pn.position = position
        for p_el in hf._element.iter(ooxml.W_P):
            if ooxml.has_page_field(p_el):
                jc = self._ppr_attr(p_el, "w:jc", "w:val")
                pn.alignment = JC_MAP.get(jc or "left", "left")
                break
        first_has = first_page and ooxml.has_page_field(
            (s.first_page_footer if position == "footer" else s.first_page_header)._element)
        pn.hide_on_first_page = bool(first_page and not first_has)
        self.report.append(
            f"Нумерация страниц: {'внизу' if position == 'footer' else 'вверху'}, "
            f"{_ALIGN_RU[pn.alignment]}"
            f"{', без номера на титульном' if pn.hide_on_first_page else ''}")

    # ============================================================== тема
    def _read_theme_fonts(self) -> dict[str, str]:
        try:
            theme = self.document.part.part_related_by(RT.THEME)
        except (KeyError, ValueError):
            return {}
        try:
            root = etree.fromstring(theme.blob)
        except etree.XMLSyntaxError:
            return {}
        out = {}
        for kind in ("major", "minor"):
            latin = root.find(f".//{{{A_NS}}}{kind}Font/{{{A_NS}}}latin")
            if latin is not None and latin.get("typeface"):
                out[kind] = latin.get("typeface")
        return out


_ALIGN_RU = {"left": "по левому краю", "center": "по центру", "right": "по правому краю",
             "justify": "по ширине"}


def _bib_format(text: str) -> str | None:
    if re.match(r"^\[\d+\]\s", text):
        return "[{n}] "
    if re.match(r"^\d+\.\s", text):
        return "{n}. "
    if re.match(r"^\d+\)\s", text):
        return "{n}) "
    return None


def _heading_format(text: str, level: int) -> str | None:
    """Определить формат номера заголовка по его тексту."""
    if level == 1:
        m = re.match(r"^(глава)\s+\d+(\.?)\s+", text, re.I)
        if m:
            return f"{m.group(1)} {{n}}{m.group(2)} {{title}}"
        m = re.match(r"^\d+(\.?)\s+\S", text)
        if m:
            return f"{{n}}{m.group(1)} {{title}}"
        return None
    nums = r"\.".join([r"\d+"] * level)
    m = re.match(rf"^{nums}(\.?)\s+\S", text)
    if m:
        keys = ["{n}", "{m}", "{k}"][:level]
        return ".".join(keys) + m.group(1) + " {title}"
    return None
