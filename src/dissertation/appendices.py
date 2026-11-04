"""Приложения как управляемый список — по аналогии со списком литературы.

Обозначение приложения: ``ПРИЛОЖЕНИЕ <bookmarkStart _GostApp_P2/>Б<bookmarkEnd/>``.
Ключ (P2) постоянный, буква — по текущему порядку (без Ё, З, Й, О, Ч, Ъ, Ы, Ь).
Ссылка в тексте ``{@P2}`` — поле REF на эту закладку: «см. приложение {@P2}» →
«см. приложение Б»; при перестановке приложений буквы в тексте обновляются.
"""
from __future__ import annotations

from dataclasses import dataclass

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import fields, ooxml
from .outline import Kind, StyleResolver
from .runs import make_run, visible_text
from .style_manager import Role
from .xref import APP_PREFIX, next_key, ref_counts, ref_order


@dataclass
class AppendixEntry:
    key: str | None
    element: object          # абзац-обозначение «ПРИЛОЖЕНИЕ А»
    letter: str
    title: str
    managed: bool            # наш стиль (буква пересчитывается)


class Appendices:
    def __init__(self, builder):
        self.builder = builder
        self.document = builder.document
        self.profile = builder.profile

    @staticmethod
    def key_of(p_el) -> str | None:
        for b in p_el.iter(fields.W_BM_START):
            name = b.get(qn("w:name")) or ""
            if name.startswith(APP_PREFIX):
                return name[len(APP_PREFIX):]
        return None

    def _title_of(self, el, resolver) -> str:
        nxt = el.getnext()
        if nxt is not None and nxt.tag == ooxml.W_P and resolver.role_of(nxt) == Role.APPENDIX_TITLE:
            return visible_text(Paragraph(nxt, self.document._body))
        return ""

    def entries(self) -> list[AppendixEntry]:
        resolver = StyleResolver(self.document)
        out = []
        for i, it in enumerate([it for it in self.builder.outline() if it.kind == Kind.APPENDIX]):
            managed = resolver.role_of(it.element) == Role.APPENDIX
            out.append(AppendixEntry(self.key_of(it.element), it.element, self.profile.appendix_letter(i),
                                     self._title_of(it.element, resolver), managed))
        return out

    def letters(self) -> dict[str, str]:
        return {e.key: e.letter for e in self.entries() if e.key}

    def citation_counts(self) -> dict[str, int]:
        return ref_counts(self.document, APP_PREFIX)

    def mention_order(self) -> list[str]:
        return ref_order(self.document, APP_PREFIX)

    # ---------------------------------------------------------------- пересчёт
    def _build_label(self, p_el, key: str, letter: str, bm_id: int) -> None:
        for child in list(p_el):
            if child.tag != qn("w:pPr"):
                p_el.remove(child)
        fmt = self.profile.appendix_label
        before, _, after = fmt.partition("{letter}")
        if before:
            p_el.append(make_run(before))
        run = make_run(letter)
        p_el.append(run)
        fields.wrap_in_bookmark(run, run, APP_PREFIX + key, bm_id)
        if after:
            p_el.append(make_run(after))

    def refresh(self) -> None:
        """Буквы наших приложений (ключи выдаются при необходимости) + кэш ссылок {@P…}."""
        bm_id = fields.next_bookmark_id(self.document)
        taken: set[str] = set()
        for e in self.entries():
            if not e.managed:
                continue
            key = e.key
            if key is None:
                key = next_key(self.document, "P", APP_PREFIX, taken)
                self._build_label(e.element, key, e.letter, bm_id)
                bm_id += 1
            else:
                start = fields.find_bookmark(e.element, APP_PREFIX + key)
                run = start.getnext() if start is not None else None
                t = run.find(qn("w:t")) if run is not None and run.tag == qn("w:r") else None
                if t is None:
                    self._build_label(e.element, key, e.letter, bm_id)
                    bm_id += 1
                elif t.text != e.letter:
                    t.text = e.letter
            taken.add(key)

    # ---------------------------------------------------------------- операции
    def sort_by_mention(self) -> int:
        """Порядок первого упоминания в тексте (ГОСТ 7.32: приложения — в порядке ссылок).

        Возвращает число переставленных приложений.
        """
        from .structure_ops import StructureOps

        order = {k: i for i, k in enumerate(self.mention_order())}
        entries = self.entries()
        target = sorted(entries, key=lambda e: (order.get(e.key, len(order)), entries.index(e)))
        if [e.element for e in target] == [e.element for e in entries]:
            return 0
        ops = StructureOps(self.builder)
        groups = [ops.section_elements(e.element) for e in target]
        anchor_parent = self.document.element.body
        first_el = ops.section_elements(entries[0].element)[0]
        marker = first_el.getprevious()
        for g in groups:
            for el in g:
                anchor_parent.remove(el)
        insert_after = marker
        for g in groups:
            for el in g:
                if insert_after is None:
                    anchor_parent.insert(0, el)
                else:
                    insert_after.addnext(el)
                insert_after = el
        self.builder.renumber()
        return sum(1 for a, b in zip(target, entries) if a.element is not b.element)
