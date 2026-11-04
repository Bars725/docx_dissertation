"""Очистка «сырого» текста перед вставкой.

Типичные источники мусора: копирование из PDF (жёсткие переносы строк внутри
абзаца, переносы слов «рас-\\nпространение», лигатуры ﬁ, разрывы страниц \\f),
из браузера (неразрывные пробелы, zero-width символы), из Word (мягкие
переносы U+00AD, табуляции вместо отступа).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

LIGATURES = {"ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi", "ﬄ": "ffl"}
INVISIBLE = dict.fromkeys(map(ord, "­​‌‍⁠﻿"), None)
SENTENCE_END = tuple(".!?…:;»\")]")
# Частицы/предлоги, перед которыми дефис — часть слова, а не перенос.
KEEP_HYPHEN_NEXT = ("то", "либо", "нибудь", "ка", "таки", "за", "под", "де")
KEEP_HYPHEN_PREV = ("по", "во", "кое", "из", "в", "к", "с")
LIST_MARKER = re.compile(r"^\s*(?:\[\d+\]|\d+[.)]|[-–—•*])\s+")
BIB_NUMBER = re.compile(r"^\s*(?:\[\d+\]|\d+\s*[.)])\s*")


@dataclass
class CleanOptions:
    paragraph_mode: str = "auto"       # auto | blank_line | line
    join_hyphenation: bool = True
    collapse_spaces: bool = True
    fix_punctuation_spaces: bool = True
    typography: bool = False           # «ёлочки» и тире
    drop_page_number_lines: bool = False


class TextCleaner:
    def __init__(self, options: CleanOptions | None = None):
        self.options = options or CleanOptions()

    # ------------------------------------------------------------- публичное
    def to_paragraphs(self, raw: str) -> list[str]:
        """Сырой текст → список абзацев."""
        text = self.normalize(raw)
        lines = text.split("\n")
        if self.options.drop_page_number_lines:
            lines = [ln for ln in lines if not re.fullmatch(r"\s*[-–]?\s*\d{1,4}\s*[-–]?\s*", ln)]
        mode = self.options.paragraph_mode
        if mode == "auto":
            has_blank = any(not ln.strip() for ln in lines[1:-1]) and len(lines) > 2
            mode = "blank_line" if has_blank else "smart"
        if mode == "line":
            paras = [ln for ln in lines if ln.strip()]
        elif mode == "blank_line":
            paras, buf = [], []
            for ln in lines + [""]:
                if ln.strip():
                    buf.append(ln)
                elif buf:
                    paras.append(self._join(buf))
                    buf = []
        else:  # smart: склеиваем строку со следующей, если это явный перенос
            paras, buf = [], []
            for ln in lines:
                if not ln.strip():
                    continue
                if buf and not self._wrap_continues(buf[-1], ln):
                    paras.append(self._join(buf))
                    buf = []
                buf.append(ln)
            if buf:
                paras.append(self._join(buf))
        return [p for p in (self._finish(p) for p in paras) if p]

    def lines(self, raw: str, keep_empty: bool = False) -> list[str]:
        """Построчно (титульный лист, список литературы): без склейки строк."""
        out = []
        for ln in self.normalize(raw).split("\n"):
            ln = self._finish(ln)
            if ln or keep_empty:
                out.append(ln)
        if keep_empty:
            while out and not out[-1]:
                out.pop()
            while out and not out[0]:
                out.pop(0)
        return out

    def bibliography_entries(self, raw: str) -> list[str]:
        """Записи списка литературы без ручной нумерации — номера проставит сборщик."""
        entries = []
        for ln in self.lines(raw):
            ln = BIB_NUMBER.sub("", ln, count=1).strip()
            if ln:
                entries.append(ln)
        return entries

    def heading(self, raw: str) -> str:
        """Заголовок: одна строка, без точки в конце (ГОСТ), без лишних пробелов."""
        text = " ".join(self.normalize(raw).split())
        text = re.sub(r"(?<!\.)\.$", "", text)  # одиночная точка; многоточие оставляем
        return text.strip()

    # ------------------------------------------------------------ внутреннее
    def normalize(self, raw: str) -> str:
        text = raw.replace("\r\n", "\n").replace("\r", "\n").replace("\f", "\n\n")
        text = text.replace(" ", "\n").replace(" ", "\n\n").replace("\v", "\n")
        text = text.translate(INVISIBLE)
        for k, v in LIGATURES.items():
            text = text.replace(k, v)
        text = text.replace(" ", " ").replace("\t", " ")
        text = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", "", text)
        return text

    def _wrap_continues(self, prev: str, nxt: str) -> bool:
        p, n = prev.rstrip(), nxt.lstrip()
        if not p or not n:
            return False
        if LIST_MARKER.match(nxt):
            return False
        if self.options.join_hyphenation and re.search(r"\w-$", p) and n[0].islower():
            return True
        if p.endswith(SENTENCE_END):
            return False
        return n[0].islower() or n[0].isdigit() or n[0] in "(«\"—–-"

    def _join(self, lines: list[str]) -> str:
        out = lines[0].strip()
        for ln in lines[1:]:
            ln = ln.strip()
            m = re.search(r"(\w+)-$", out)
            if self.options.join_hyphenation and m and ln and ln[0].islower():
                word_before = m.group(1).lower()
                next_word = re.match(r"\w+", ln)
                next_word = next_word.group(0).lower() if next_word else ""
                if next_word in KEEP_HYPHEN_NEXT or word_before in KEEP_HYPHEN_PREV:
                    out = out + ln            # «из-за», «что-то»: дефис сохраняем
                else:
                    out = out[:-1] + ln       # «рас-пространение» → «распространение»
            else:
                out = f"{out} {ln}"
        return out

    def _finish(self, text: str) -> str:
        o = self.options
        if o.collapse_spaces:
            text = re.sub(r" {2,}", " ", text).strip()
        if o.fix_punctuation_spaces:
            text = re.sub(r" +([,.;:!?%)\]»])", r"\1", text)
            text = re.sub(r"([(\[«]) +", r"\1", text)
        if o.typography:
            text = self._typography(text)
        return text

    @staticmethod
    def _typography(text: str) -> str:
        text = re.sub(r"(^|[\s(\[])\"", r"\1«", text)
        text = text.replace('"', "»")
        text = re.sub(r" -{1,2} ", " – ", text)
        text = re.sub(r"^-{1,2} ", "– ", text)
        return text
