"""Перекрёстные ссылки в тексте: на источники и на приложения.

    [@R3]  [@R3, с. 15]  [@R1; @R4]   → [3] [3, с. 15] [1; 4]   (номер записи списка литературы)
    {@P2}                             → Б                       (буква приложения)
    {@F3}                             → 2  (или 1.2, А.1)       (номер рисунка)

Каждая ссылка — поле Word ``REF <закладка> \\h``; программа пересчитывает кэш полей,
Word — по F9. Ключи (R3, P2) постоянные: при сортировке меняются только номера/буквы.
"""
from __future__ import annotations

import re

from . import fields
from .runs import make_run

BIB_PREFIX = "_GostRef_"
APP_PREFIX = "_GostApp_"
FIG_PREFIX = "_GostFig_"
CITE_RX = re.compile(r"\[@([^\[\]]+)\]")
APP_RX = re.compile(r"\{@([A-Za-z][A-Za-z0-9]*)\}")
KEY_RX = re.compile(r"^[A-Za-z][A-Za-z0-9]*$")


def parse_citation(inner: str):
    """«R3, с. 15; @R5» → [(R3, «с. 15»), (R5, «»)]; None — если это не ссылка."""
    parts = []
    for chunk in inner.split(";"):
        chunk = chunk.strip().lstrip("@").strip()
        key, _, loc = chunk.partition(",")
        key = key.strip()
        if not KEY_RX.match(key):
            return None
        parts.append((key, loc.strip()))
    return parts or None


def _tokens(text: str):
    """[(start, end, kind, payload)] по порядку; kind: 'cite' | 'app'."""
    out = []
    for m in CITE_RX.finditer(text):
        parts = parse_citation(m.group(1))
        if parts:
            out.append((m.start(), m.end(), "cite", parts))
    for m in APP_RX.finditer(text):
        out.append((m.start(), m.end(), "app", m.group(1)))
    out.sort()
    result, last_end = [], -1
    for tok in out:
        if tok[0] >= last_end:
            result.append(tok)
            last_end = tok[1]
    return result


def has_refs(text: str) -> bool:
    return bool(text) and ("@" in text) and bool(_tokens(text))


class XRef:
    def __init__(self, builder):
        self.builder = builder
        self.sep = builder.profile.citation_separator

    def _numbers(self) -> dict[str, int]:
        return self.builder.bibliography.numbers()

    def _letters(self) -> dict[str, str]:
        """Ключи {@…}: приложения (P…) и рисунки (F…) → буква / номер."""
        from .figures import Figures

        out = dict(self.builder.appendices.letters())
        out.update(Figures(self.builder).numbers())
        return out

    @staticmethod
    def _prefix(key: str) -> str:
        return FIG_PREFIX if key.startswith("F") else APP_PREFIX

    def rich_runs(self, text: str, unknown: list[str] | None = None) -> list:
        numbers, letters = self._numbers(), self._letters()
        runs, pos = [], 0
        for start, end, kind, payload in _tokens(text):
            if start > pos:
                runs.append(make_run(text[pos:start]))
            if kind == "cite":
                runs.append(make_run("["))
                for i, (key, loc) in enumerate(payload):
                    if i:
                        runs.append(make_run(self.sep))
                    num = numbers.get(key)
                    if num is None and unknown is not None:
                        unknown.append(key)
                    runs += fields.field_runs(f"REF {BIB_PREFIX}{key} \\h", str(num) if num else "?")
                    if loc:
                        runs.append(make_run(", " + loc))
                runs.append(make_run("]"))
            else:
                letter = letters.get(payload)
                if letter is None and unknown is not None:
                    unknown.append(payload)
                runs += fields.field_runs(f"REF {self._prefix(payload)}{payload} \\h", letter or "?")
            pos = end
        if pos < len(text):
            runs.append(make_run(text[pos:]))
        return runs

    def render(self, text: str) -> str:
        """Для предпросмотра: токены → текущие номера/буквы."""
        if not has_refs(text):
            return text
        numbers, letters = self._numbers(), self._letters()
        out, pos = [], 0
        for start, end, kind, payload in _tokens(text):
            out.append(text[pos:start])
            if kind == "cite":
                items = [(str(numbers[k]) if k in numbers else "?") + (f", {loc}" if loc else "")
                         for k, loc in payload]
                out.append("[" + self.sep.join(items) + "]")
            else:
                out.append(letters.get(payload, "?"))
            pos = end
        out.append(text[pos:])
        return "".join(out)

    def refresh_fields(self) -> None:
        numbers, letters = self._numbers(), self._letters()
        for f in fields.iter_complex_fields(self.builder.document.element.body):
            target = fields.ref_target(f.instr) or ""
            if target.startswith(BIB_PREFIX):
                n = numbers.get(target[len(BIB_PREFIX):])
                fields.set_field_result(f, str(n) if n else "?")
            elif target.startswith(APP_PREFIX):
                fields.set_field_result(f, letters.get(target[len(APP_PREFIX):], "?"))
            elif target.startswith(FIG_PREFIX):
                fields.set_field_result(f, letters.get(target[len(FIG_PREFIX):], "?"))


def ref_counts(document, prefix: str) -> dict[str, int]:
    out: dict[str, int] = {}
    for f in fields.iter_complex_fields(document.element.body):
        target = fields.ref_target(f.instr) or ""
        if target.startswith(prefix):
            key = target[len(prefix):]
            out[key] = out.get(key, 0) + 1
    return out


def ref_order(document, prefix: str) -> list[str]:
    seen = []
    for f in fields.iter_complex_fields(document.element.body):
        target = fields.ref_target(f.instr) or ""
        if target.startswith(prefix) and target[len(prefix):] not in seen:
            seen.append(target[len(prefix):])
    return seen


def next_key(document, letter: str, prefix: str, taken=()) -> str:
    """Новый ключ больше всех занятых: в закладках, в ссылках текста и в taken."""
    nums = []
    for name in fields.bookmark_names(document.element.body, prefix):
        tail = name[len(prefix):]
        if tail[:1] == letter and tail[1:].isdigit():
            nums.append(int(tail[1:]))
    for key in list(ref_counts(document, prefix)) + list(taken):
        if key and key[:1] == letter and key[1:].isdigit():
            nums.append(int(key[1:]))
    return f"{letter}{max(nums, default=0) + 1}"
