"""Профиль оформления — независимое от .docx описание всех параметров ГОСТ.

Профиль хранится в JSON и не содержит ссылок на XML конкретного документа:
его можно извлечь из одного файла и применить к другому.
"""
from __future__ import annotations

import copy
import json
from dataclasses import asdict, dataclass, field, fields, is_dataclass
from pathlib import Path
from typing import Any

PROFILE_VERSION = 1

ALIGNMENTS = ("left", "center", "right", "justify")


@dataclass
class PageSetup:
    """Параметры страницы (мм)."""

    width_mm: float = 210.0
    height_mm: float = 297.0
    margin_left_mm: float = 30.0
    margin_right_mm: float = 15.0
    margin_top_mm: float = 20.0
    margin_bottom_mm: float = 20.0
    header_distance_mm: float = 12.5
    footer_distance_mm: float = 12.5


@dataclass
class ParagraphSpec:
    """Оформление одного вида абзацев (переносится в стиль Word)."""

    font_name: str = "Times New Roman"
    font_size_pt: float = 14.0
    bold: bool = False
    italic: bool = False
    all_caps: bool = False
    alignment: str = "justify"          # left | center | right | justify
    line_spacing: float = 1.5           # множитель (lineRule="auto")
    first_line_indent_cm: float = 1.25
    left_indent_cm: float = 0.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    page_break_before: bool = False
    keep_with_next: bool = False
    no_hyphenation: bool = False        # w:suppressAutoHyphens (для заголовков)


@dataclass
class PageNumbering:
    enabled: bool = True
    position: str = "footer"            # footer | header
    alignment: str = "center"
    hide_on_first_page: bool = True     # титульный лист считается, но номер не ставится


@dataclass
class HeadingNumbering:
    """Текстовая нумерация заголовков. {n}, {m}, {k} — номера уровней, {title} — текст.

    enabled=False — номера не добавляются (например, если в образце нумерация
    сделана многоуровневым списком Word или её нет вовсе).
    """

    enabled: bool = True
    chapter: str = "{n} {title}"
    subchapter: str = "{n}.{m} {title}"
    section: str = "{n}.{m}.{k} {title}"


@dataclass
class FormattingProfile:
    name: str = "ГОСТ 7.32-2017"
    source: str = "builtin"
    version: int = PROFILE_VERSION

    page: PageSetup = field(default_factory=PageSetup)

    body: ParagraphSpec = field(default_factory=ParagraphSpec)
    title_page: ParagraphSpec = field(default_factory=ParagraphSpec)
    structural_heading: ParagraphSpec = field(default_factory=ParagraphSpec)
    chapter_heading: ParagraphSpec = field(default_factory=ParagraphSpec)
    subchapter_heading: ParagraphSpec = field(default_factory=ParagraphSpec)
    section_heading: ParagraphSpec = field(default_factory=ParagraphSpec)
    bibliography: ParagraphSpec = field(default_factory=ParagraphSpec)
    appendix_heading: ParagraphSpec = field(default_factory=ParagraphSpec)
    appendix_title: ParagraphSpec = field(default_factory=ParagraphSpec)

    numbering: HeadingNumbering = field(default_factory=HeadingNumbering)
    page_numbering: PageNumbering = field(default_factory=PageNumbering)

    structural_uppercase: bool = True
    introduction_title: str = "ВВЕДЕНИЕ"
    conclusion_title: str = "ЗАКЛЮЧЕНИЕ"
    bibliography_title: str = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
    bibliography_number_format: str = "{n}. "
    appendix_label: str = "ПРИЛОЖЕНИЕ {letter}"
    # ГОСТ: без Ё, З, Й, О, Ч, Ъ, Ы, Ь
    appendix_letters: str = "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ"

    language: str = "ru-RU"
    auto_hyphenation: bool = False

    # Что удалось извлечь из образца, а что взято из ГОСТ по умолчанию.
    extraction_report: list[str] = field(default_factory=list)

    # ------------------------------------------------------------------ JSON
    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

    def save(self, path: str | Path) -> None:
        Path(path).write_text(self.to_json(), encoding="utf-8")

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "FormattingProfile":
        obj = cls()
        _update_dataclass(obj, data)
        return obj

    @classmethod
    def load(cls, path: str | Path) -> "FormattingProfile":
        return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))

    def copy(self) -> "FormattingProfile":
        return copy.deepcopy(self)

    def appendix_letter(self, index: int) -> str:
        """Буква приложения по порядковому номеру (0 → А). При нехватке букв — цифры."""
        letters = self.appendix_letters
        return letters[index] if index < len(letters) else str(index + 1)


def _update_dataclass(obj: Any, data: dict[str, Any]) -> None:
    """Мягкое обновление: неизвестные ключи игнорируются, отсутствующие остаются по умолчанию.

    Это даёт совместимость старых JSON-профилей с новыми версиями программы.
    """
    for f in fields(obj):
        if f.name not in data:
            continue
        current = getattr(obj, f.name)
        value = data[f.name]
        if is_dataclass(current) and isinstance(value, dict):
            _update_dataclass(current, value)
        elif isinstance(current, bool):
            setattr(obj, f.name, bool(value))
        elif isinstance(current, float) and isinstance(value, (int, float)):
            setattr(obj, f.name, float(value))
        else:
            setattr(obj, f.name, value)
