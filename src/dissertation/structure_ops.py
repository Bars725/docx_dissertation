"""Операции над структурой: чтение/правка содержимого пункта, перемещение, удаление.

«Раздел» пункта — его заголовок, собственный текст и все вложенные подпункты
(до следующего заголовка того же или более высокого уровня). «Собственный текст» —
абзацы между заголовком и первым следующим заголовком любого уровня.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import fields, ooxml
from .bibliography import BM_PREFIX, has_citations
from .outline import Kind, OutlineItem, StyleResolver
from .runs import make_run, set_paragraph_text, visible_text
from .style_manager import Role
from .text_cleaner import CleanOptions, TextCleaner

MOVABLE = {Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION, Kind.APPENDIX}
KIND_RU = {Kind.CHAPTER: "глава", Kind.SUBCHAPTER: "подглава", Kind.SECTION: "пункт",
           Kind.APPENDIX: "приложение", Kind.INTRODUCTION: "введение", Kind.CONCLUSION: "заключение",
           Kind.BIBLIOGRAPHY: "список литературы", Kind.TOC: "оглавление",
           Kind.TITLE: "титульный лист", Kind.OTHER_STRUCTURAL: "структурный элемент"}
_REF_MARK = re.compile(r"\x01([A-Za-z][A-Za-z0-9]*)\x02")


@dataclass
class SectionContent:
    kind: Kind
    heading: str = ""
    heading_editable: bool = True
    body: str = ""
    body_editable: bool = True
    reason: str = ""                     # почему текст только для чтения
    paragraphs: int = 0


@dataclass
class OpResult:
    ok: bool
    message: str
    warnings: list[str] = field(default_factory=list)


class StructureOps:
    def __init__(self, builder):
        self.b = builder
        self.doc = builder.document
        self.body_el = builder.document.element.body

    # ========================================================= навигация
    def _item(self, element) -> tuple[int, list[OutlineItem]]:
        outline = self.b.outline()
        for i, it in enumerate(outline):
            if it.element is element:
                return i, outline
        raise ValueError("Элемент структуры не найден — обновите структуру")

    def _top(self, el):
        while el is not None and el.getparent() is not self.body_el:
            el = el.getparent()
        return el

    def _children(self) -> list:
        return [c for c in self.body_el if c.tag != ooxml.W_SECTPR]

    def section_elements(self, element) -> list:
        """Заголовок + собственный текст + подпункты (элементы верхнего уровня тела)."""
        i, outline = self._item(element)
        item = outline[i]
        if item.kind == Kind.TITLE:
            return [self._top(it.element) for it in outline if it.kind == Kind.TITLE]
        children = self._children()
        start = children.index(self._top(item.element))
        end = len(children)
        for later in outline[i + 1:]:
            if later.kind == Kind.TITLE:
                continue
            if later.level <= item.level:
                end = children.index(self._top(later.element))
                break
        return children[start:end]

    def own_body(self, element) -> list:
        """Собственные абзацы пункта (без заголовка, заголовка приложения и подпунктов)."""
        i, outline = self._item(element)
        children = self._children()
        start = children.index(self._top(outline[i].element)) + 1
        end = len(children)
        for later in outline[i + 1:]:
            end = children.index(self._top(later.element))
            break
        els = children[start:end]
        resolver = StyleResolver(self.doc)
        if outline[i].kind == Kind.APPENDIX and els and els[0].tag == ooxml.W_P \
                and resolver.role_of(els[0]) == Role.APPENDIX_TITLE:
            els = els[1:]
        return els

    def _appendix_title(self, element):
        nxt = self._top(element).getnext()
        if nxt is not None and nxt.tag == ooxml.W_P and \
                StyleResolver(self.doc).role_of(nxt) == Role.APPENDIX_TITLE:
            return nxt
        return None

    def parent_and_siblings(self, element):
        i, outline = self._item(element)
        item = outline[i]

        def parent_of(j):
            for k in range(j - 1, -1, -1):
                if outline[k].level < outline[j].level and outline[k].kind != Kind.TITLE:
                    return k
            return None

        p = parent_of(i)
        sibs = [j for j, it in enumerate(outline)
                if it.level == item.level and it.kind != Kind.TITLE and parent_of(j) == p]
        return outline, i, sibs

    # ========================================================= чтение
    def read(self, element) -> SectionContent:
        i, outline = self._item(element)
        item = outline[i]
        content = SectionContent(kind=item.kind)
        para = Paragraph(item.element, self.doc._body) if item.element.tag == ooxml.W_P else None
        if item.kind in (Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION):
            content.heading = self._strip_number(item)
        elif item.kind == Kind.APPENDIX:
            t = self._appendix_title(item.element)
            content.heading = visible_text(Paragraph(t, self.doc._body)) if t is not None else ""
        else:
            content.heading = visible_text(para) if para is not None else item.text
            content.heading_editable = False
        texts, reasons = [], []
        for el in self.own_body(element):
            reason = self._complex(el)
            if reason:
                reasons.append(reason)
            if el.tag == ooxml.W_P:
                t = self._editable_text(el)
                if t.strip():
                    texts.append(t)
        content.body = "\n\n".join(texts)
        content.paragraphs = len(texts)
        if reasons:
            content.body_editable = False
            content.reason = "в тексте есть: " + ", ".join(sorted(set(reasons))) + \
                             " — текст можно изменить только в Word, заголовок — здесь"
        return content

    def _strip_number(self, item: OutlineItem) -> str:
        from .document_builder import _format_regex

        num = self.b.profile.numbering
        fmt = {Kind.CHAPTER: num.chapter, Kind.SUBCHAPTER: num.subchapter,
               Kind.SECTION: num.section}[item.kind]
        text = " ".join(ooxml.element_text(item.element).split())
        if num.enabled:
            m = _format_regex(fmt).match(text)
            if m:
                return m.group("title").strip()
        return text

    @staticmethod
    def _complex(el) -> str | None:
        if el.tag == ooxml.W_TBL:
            return "таблица"
        if el.tag != ooxml.W_P:
            return "встроенный элемент"
        for tag, name in (("w:drawing", "рисунок"), ("w:pict", "рисунок"), ("w:object", "объект"),
                          ("w:footnoteReference", "сноска"), ("w:endnoteReference", "сноска"),
                          ("w:txbxContent", "надпись")):
            if el.find(".//" + qn(tag)) is not None:
                return name
        if el.find(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath") is not None:
            return "формула"
        ppr = el.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:numPr")) is not None:
            return "список"
        for f in fields.iter_complex_fields(el):
            target = fields.ref_target(f.instr)
            if not (target and target.startswith(BM_PREFIX)):
                return "поле Word"
        return None

    def _editable_text(self, p_el) -> str:
        """Текст абзаца; поля-ссылки на литературу → @КЛЮЧ (получается [@R3, с. 15])."""
        out = []
        stack = []
        for r in p_el.iter(qn("w:r")):
            for child in r:
                if child.tag == fields.W_FLDCHAR:
                    kind = child.get(qn("w:fldCharType"))
                    if kind == "begin":
                        stack.append({"instr": "", "state": "instr"})
                    elif kind == "separate" and stack:
                        stack[-1]["state"] = "result"
                    elif kind == "end" and stack:
                        f = stack.pop()
                        target = fields.ref_target(f["instr"])
                        if target and target.startswith(BM_PREFIX):
                            out.append(f"\x01{target[len(BM_PREFIX):]}\x02")
                elif child.tag == fields.W_INSTR and stack:
                    stack[-1]["instr"] += child.text or ""
                elif stack:
                    continue          # результат поля пропускаем — восстановим токеном
                elif child.tag == qn("w:t"):
                    out.append(child.text or "")
                elif child.tag == qn("w:tab"):
                    out.append("\t")
                elif child.tag in (qn("w:br"), qn("w:cr")):
                    out.append("\n")
        return _REF_MARK.sub(lambda m: "@" + m.group(1), "".join(out))

    # ========================================================= правка
    def update(self, element, heading: str | None, body: str | None, clean: bool = True,
               options: CleanOptions | None = None) -> OpResult:
        content = self.read(element)
        i, outline = self._item(element)
        item = outline[i]
        cleaner = TextCleaner(options or CleanOptions())
        warnings: list[str] = []
        if heading is not None and content.heading_editable:
            title = cleaner.heading(heading)
            if not title:
                return OpResult(False, "Заголовок не может быть пустым")
            if item.kind == Kind.APPENDIX:
                t = self._appendix_title(item.element)
                if t is not None:
                    set_paragraph_text(Paragraph(t, self.doc._body), title)
                else:
                    add = self.b._inserter("after", self._top(item.element))
                    add(Role.APPENDIX_TITLE, title)
            else:
                from .document_builder import BlockType

                block = {Kind.CHAPTER: BlockType.CHAPTER, Kind.SUBCHAPTER: BlockType.SUBCHAPTER,
                         Kind.SECTION: BlockType.SECTION}[item.kind]
                set_paragraph_text(Paragraph(item.element, self.doc._body),
                                   self.b._numbered_heading(block, title))
        if body is not None and content.body_editable:
            paras = cleaner.to_paragraphs(body) if clean else \
                [ln for ln in body.split("\n") if ln.strip()]
            old = self.own_body(element)
            anchor = self._appendix_title(item.element) if item.kind == Kind.APPENDIX else None
            anchor = anchor if anchor is not None else self._top(item.element)
            for el in old:
                el.getparent().remove(el)
            add = self.b._inserter("after", anchor)
            self.b._unknown_refs = []
            for t in paras:
                add(Role.BODY, t)
            if self.b._unknown_refs:
                warnings.append("Нет в списке литературы: " + ", ".join(sorted(set(self.b._unknown_refs))))
        elif body is not None and not content.body_editable and body.strip() != content.body.strip():
            warnings.append("Текст не изменён: " + content.reason)
        self.b.renumber()
        return OpResult(True, "Изменения сохранены", warnings)

    def preview_update(self, element, heading: str, body: str, clean: bool = True,
                       options: CleanOptions | None = None):
        """Предпросмотр правки: тот же номер пункта, новый заголовок и текст."""
        from .document_builder import PreviewResult

        content = self.read(element)
        i, outline = self._item(element)
        item = outline[i]
        cleaner = TextCleaner(options or CleanOptions())
        res = PreviewResult(where=f"вместо текущего содержимого «{item.text}»")
        role = {Kind.CHAPTER: Role.CHAPTER, Kind.SUBCHAPTER: Role.SUBCHAPTER,
                Kind.SECTION: Role.SECTION, Kind.APPENDIX: Role.APPENDIX}.get(item.kind, Role.STRUCTURAL)
        title = cleaner.heading(heading) if heading.strip() else ""
        if item.kind in (Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION) and content.heading_editable:
            full = " ".join(ooxml.element_text(item.element).split())
            old_title = content.heading
            prefix = full[: len(full) - len(old_title)] if full.endswith(old_title) else ""
            res.items.append((role, prefix + (title or old_title)))
        else:
            res.items.append((role, item.text))
            if item.kind == Kind.APPENDIX and (title or content.heading):
                res.items.append((Role.APPENDIX_TITLE, title or content.heading))
        if content.body_editable:
            paras = cleaner.to_paragraphs(body) if clean else [ln for ln in body.split("\n") if ln.strip()]
            bib = self.b.bibliography
            paras = [bib.render_citations(p) if has_citations(p) else p for p in paras]
        else:
            paras = [p for p in content.body.split("\n\n") if p]
            res.warnings.append("Только для чтения: " + content.reason)
        res.items += [(Role.BODY, p) for p in paras]
        res.page_breaks = [self.b._role_breaks(r) for r, _ in res.items]
        return res

    # ======================================================= перемещение
    def move(self, element, delta: int) -> OpResult:
        outline, i, sibs = self.parent_and_siblings(element)
        item = outline[i]
        if item.kind not in MOVABLE:
            return OpResult(False, f"«{item.text}» — {KIND_RU.get(item.kind, 'элемент')} с "
                                   "фиксированным местом по ГОСТ, переместить нельзя")
        pos = sibs.index(i)
        j = pos + (1 if delta > 0 else -1)
        if j < 0 or j >= len(sibs):
            return OpResult(False, "Дальше перемещать некуда")
        neighbor = outline[sibs[j]]
        if neighbor.kind != item.kind:
            where = "выше" if delta < 0 else "ниже"
            return OpResult(False, f"{KIND_RU[item.kind].capitalize()} не может стоять {where} — "
                                   f"там «{neighbor.text}»")
        mine = self.section_elements(item.element)
        theirs = self.section_elements(neighbor.element)
        if delta < 0:
            target = theirs[0]
            for el in mine:
                target.addprevious(el)
        else:
            target = theirs[-1]
            for el in reversed(mine):
                target.addnext(el)
        self.b.renumber()
        return OpResult(True, "Перемещено, номера пересчитаны")

    # ========================================================= удаление
    def describe(self, element) -> str:
        i, outline = self._item(element)
        item = outline[i]
        els = self.section_elements(element)
        subs = sum(1 for it in outline[i + 1:] if it.level > item.level
                   and self._top(it.element) in els)
        paras = sum(1 for el in els[1:] if el.tag == ooxml.W_P and ooxml.element_text(el).strip())
        tables = sum(1 for el in els if el.tag == ooxml.W_TBL)
        parts = [f"абзацев: {paras}"]
        if subs:
            parts.append(f"подпунктов: {subs}")
        if tables and item.kind != Kind.TITLE:
            parts.append(f"таблиц: {tables}")
        return f"«{item.text}» ({KIND_RU.get(item.kind, 'элемент')}; " + ", ".join(parts) + ")"

    def delete(self, element) -> OpResult:
        i, outline = self._item(element)
        item = outline[i]
        refs_before = sum(self.b.bibliography.citation_counts().values()) \
            if item.kind == Kind.BIBLIOGRAPHY else 0
        for el in self.section_elements(element):
            el.getparent().remove(el)
        self.b.renumber()
        warnings = []
        if refs_before:
            warnings.append(f"Ссылок на литературу в тексте: {refs_before} — они показывают «[?]»")
        return OpResult(True, f"Удалено: {item.text}", warnings)


__all__ = ["StructureOps", "SectionContent", "OpResult", "has_citations", "make_run"]
