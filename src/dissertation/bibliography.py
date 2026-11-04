"""Список литературы как набор записей с постоянными ключами и «живые» ссылки на них.

Устройство записи в .docx (стиль «ГОСТ Список литературы»):

    [префикс] <bookmarkStart _GostRef_R3/> <r>3</r> <bookmarkEnd/> <r>. Иванов И. И. …</r>

Номер записи лежит внутри закладки ``_GostRef_<ключ>``. Ссылка в тексте главы —
поле ``REF _GostRef_R3 \\h``: Word по F9 подставит текущий номер из закладки,
а программа при каждом изменении списка пересчитывает и номера, и кэш полей.
Поэтому после сортировки/вставки источника ссылки «[3]» остаются верными.

Синтаксис ссылок в окне ввода:  [@R3]   [@R3, с. 15]   [@R3; @R7, с. 2–4]
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import fields, ooxml
from .outline import Kind, StyleResolver, classify_paragraph
from .runs import make_run, visible_text
from .style_manager import Role
from .text_cleaner import BIB_NUMBER

BM_PREFIX = "_GostRef_"
KEY_RX = re.compile(r"^[A-Za-z][A-Za-z0-9]*$")
CITE_RX = re.compile(r"\[@([^\[\]]+)\]")


@dataclass
class BibEntry:
    key: str | None
    text: str

    @property
    def is_linked(self) -> bool:
        return self.key is not None


@dataclass
class CitationPart:
    key: str
    locator: str = ""


def parse_citation(inner: str) -> list[CitationPart] | None:
    """«R3, с. 15; @R5» → [(R3, «с. 15»), (R5, «»)]; None — если это не ссылка."""
    parts = []
    for chunk in inner.split(";"):
        chunk = chunk.strip().lstrip("@").strip()
        key, _, loc = chunk.partition(",")
        key = key.strip()
        if not KEY_RX.match(key):
            return None
        parts.append(CitationPart(key, loc.strip()))
    return parts or None


def has_citations(text: str) -> bool:
    return any(parse_citation(m.group(1)) for m in CITE_RX.finditer(text))


def alpha_key(text: str):
    """Сначала источники на кириллице, затем на латинице (ГОСТ Р 7.0.100), без учёта регистра."""
    first = next((ch for ch in text if ch.isalpha()), "")
    cyr = "а" <= first.lower() <= "я" or first.lower() == "ё"
    return (0 if cyr else 1, text.casefold().replace("ё", "е"))


class Bibliography:
    def __init__(self, builder):
        self.builder = builder
        self.document = builder.document
        self.profile = builder.profile

    # ============================================================ чтение
    def heading(self):
        return self.builder.find(Kind.BIBLIOGRAPHY)

    def _section(self) -> list:
        """Абзацы раздела списка литературы (до следующего заголовка 1 уровня)."""
        head = self.heading()
        if head is None:
            return []
        resolver = StyleResolver(self.document)
        out, started = [], False
        for el in ooxml.iter_body_blocks(self.document.element.body):
            if el is head.element:
                started = True
                continue
            if not started:
                continue
            if el.tag == ooxml.W_P:
                item = classify_paragraph(el, resolver)
                if item is not None and item.level == 0:
                    break
                out.append(el)
        return out

    def _entry_paragraphs(self) -> list:
        return [el for el in self._section() if ooxml.element_text(el).strip()]

    @staticmethod
    def _key_of(p_el) -> str | None:
        for b in p_el.iter(fields.W_BM_START):
            name = b.get(qn("w:name")) or ""
            if name.startswith(BM_PREFIX):
                return name[len(BM_PREFIX):]
        return None

    def _strip_number(self, text: str) -> str:
        fmt = self.profile.bibliography_number_format
        if "{n}" in fmt:
            before, after = fmt.split("{n}", 1)
            rx = re.compile("^" + re.escape(before) + r"\d+" + re.escape(after.rstrip()) + r"\s*")
            text2 = rx.sub("", text, count=1)
            if text2 != text:
                return text2.strip()
        return BIB_NUMBER.sub("", text, count=1).strip()

    def entries(self) -> list[BibEntry]:
        out = []
        for el in self._entry_paragraphs():
            text = visible_text(Paragraph(el, self.document._body))
            out.append(BibEntry(self._key_of(el), self._strip_number(text)))
        return out

    def numbers(self) -> dict[str, int]:
        return {e.key: i for i, e in enumerate(self.entries(), 1) if e.key}

    def citation_counts(self) -> Counter:
        c = Counter()
        for f in fields.iter_complex_fields(self.document.element.body):
            target = fields.ref_target(f.instr)
            if target and target.startswith(BM_PREFIX):
                c[target[len(BM_PREFIX):]] += 1
        return c

    def citation_order(self) -> list[str]:
        seen = []
        for f in fields.iter_complex_fields(self.document.element.body):
            target = fields.ref_target(f.instr)
            if target and target.startswith(BM_PREFIX):
                key = target[len(BM_PREFIX):]
                if key not in seen:
                    seen.append(key)
        return seen

    def _next_key(self, taken: set[str]) -> str:
        """Новый ключ больше всех занятых — включая ключи, на которые ещё ссылается текст,
        чтобы «осиротевшая» ссылка на удалённый источник не прицепилась к новому."""
        taken = set(taken) | set(self.citation_counts())
        nums = [int(k[1:]) for k in taken if k and k[0] == "R" and k[1:].isdigit()]
        names = set(fields.bookmark_names(self.document.element.body, BM_PREFIX + "R"))
        nums += [int(n[len(BM_PREFIX) + 1:]) for n in names if n[len(BM_PREFIX) + 1:].isdigit()]
        return f"R{max(nums, default=0) + 1}"

    # ============================================================ запись
    def set_entries(self, entries: list[BibEntry]) -> list[BibEntry]:
        """Переписать раздел целиком в заданном порядке; записям без ключа выдаются ключи."""
        head = self._ensure_heading()
        for el in self._section():
            el.getparent().remove(el)
        taken = {e.key for e in entries if e.key}
        result = []
        for e in entries:
            key = e.key or self._next_key(taken)
            taken.add(key)
            result.append(BibEntry(key, e.text.strip()))
        cursor = head.element
        style = self.builder.styles.style_for(Role.BIBLIOGRAPHY)
        bm_id = fields.next_bookmark_id(self.document)
        for i, e in enumerate(result, 1):
            p = ooxml.new_paragraph_after(cursor, self.document._body)
            p.style = style
            self._build_entry(p._p, e, i, bm_id)
            bm_id += 1
            cursor = p._p
        self.builder.renumber()
        return result

    def _build_entry(self, p_el, entry: BibEntry, number: int, bm_id: int) -> None:
        fmt = self.profile.bibliography_number_format
        if "{n}" not in fmt:
            fmt = "{n}. "
        before, after = fmt.split("{n}", 1)
        if before:
            p_el.append(make_run(before))
        num = make_run(str(number))
        p_el.append(num)
        fields.wrap_in_bookmark(num, num, BM_PREFIX + entry.key, bm_id)
        p_el.append(make_run(after + entry.text))

    def _ensure_heading(self):
        head = self.heading()
        if head is not None:
            return head
        from .document_builder import BlockType, InsertPosition, InsertResult

        where, anchor = self.builder._resolve_position(BlockType.BIBLIOGRAPHY, InsertPosition(),
                                                       InsertResult())
        add = self.builder._inserter(where, anchor)
        add(Role.STRUCTURAL, self.builder._structural_text(self.profile.bibliography_title))
        return self.heading()

    # ----------------------------------------------------------- операции
    def add(self, texts: list[str]) -> list[BibEntry]:
        texts = [t.strip() for t in texts if t.strip()]
        if not texts:
            return []
        current = self.entries()
        new = [BibEntry(None, t) for t in texts]
        result = self.set_entries(current + new)
        return result[len(current):]

    def update_text(self, key: str, text: str) -> None:
        self.set_entries([BibEntry(e.key, text if e.key == key else e.text) for e in self.entries()])

    def remove(self, key: str) -> int:
        """Удалить запись. Возвращает число ссылок на неё в тексте (они покажут «?»)."""
        refs = self.citation_counts().get(key, 0)
        self.set_entries([e for e in self.entries() if e.key != key])
        return refs

    def move(self, key: str, delta: int) -> None:
        items = self.entries()
        idx = next(i for i, e in enumerate(items) if e.key == key)
        new = max(0, min(len(items) - 1, idx + delta))
        items.insert(new, items.pop(idx))
        self.set_entries(items)

    def sort_alphabetically(self) -> None:
        self.set_entries(sorted(self.entries(), key=lambda e: alpha_key(e.text)))

    def sort_by_citation(self) -> None:
        """Порядок первого упоминания в тексте (ГОСТ 7.32); неупомянутые — в конце."""
        order = {k: i for i, k in enumerate(self.citation_order())}
        items = self.entries()
        self.set_entries(sorted(items, key=lambda e: (order.get(e.key, len(order)),
                                                      items.index(e))))

    # ======================================================== пересчёт
    def refresh(self) -> None:
        """Номера записей (внутри закладок) и кэш всех полей-ссылок на них."""
        numbers = {}
        for i, el in enumerate(self._entry_paragraphs(), 1):
            key = self._key_of(el)
            if key is None:
                continue
            numbers[key] = i
            start = fields.find_bookmark(el, BM_PREFIX + key)
            run = start.getnext() if start is not None else None
            if run is not None and run.tag == qn("w:r"):
                t = run.find(qn("w:t"))
                if t is not None and t.text != str(i):
                    t.text = str(i)
        for f in fields.iter_complex_fields(self.document.element.body):
            target = fields.ref_target(f.instr)
            if target and target.startswith(BM_PREFIX):
                value = numbers.get(target[len(BM_PREFIX):])
                fields.set_field_result(f, str(value) if value else "?")

    # ================================================== ссылки в тексте
    def rich_runs(self, text: str, unknown: list[str] | None = None) -> list:
        """Текст абзаца → runs; токены [@R3, с. 15] превращаются в поля REF."""
        numbers = self.numbers()
        sep = self.profile.citation_separator
        runs, pos = [], 0
        for m in CITE_RX.finditer(text):
            parts = parse_citation(m.group(1))
            if parts is None:
                continue
            if m.start() > pos:
                runs.append(make_run(text[pos:m.start()]))
            runs.append(make_run("["))
            for i, part in enumerate(parts):
                if i:
                    runs.append(make_run(sep))
                num = numbers.get(part.key)
                if num is None and unknown is not None:
                    unknown.append(part.key)
                runs += fields.field_runs(f"REF {BM_PREFIX}{part.key} \\h", str(num) if num else "?")
                if part.locator:
                    runs.append(make_run(", " + part.locator))
            runs.append(make_run("]"))
            pos = m.end()
        if pos < len(text):
            runs.append(make_run(text[pos:]))
        return runs

    def render_citations(self, text: str) -> str:
        """Для предпросмотра: [@R3, с. 15] → [3, с. 15]."""
        numbers = self.numbers()
        sep = self.profile.citation_separator

        def repl(m):
            parts = parse_citation(m.group(1))
            if parts is None:
                return m.group(0)
            items = []
            for p in parts:
                n = numbers.get(p.key)
                items.append((str(n) if n else "?") + (f", {p.locator}" if p.locator else ""))
            return "[" + sep.join(items) + "]"

        return CITE_RX.sub(repl, text)
