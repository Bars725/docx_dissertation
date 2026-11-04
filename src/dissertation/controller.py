"""Состояние приложения без привязки к GUI — тестируется без Tkinter."""
from __future__ import annotations

import re
from pathlib import Path

from .document_builder import BlockType, DocumentBuilder, InsertPosition, InsertResult
from .gost import DEFAULT_PRESET, PRESETS
from .profile import FormattingProfile
from .template_parser import TemplateParser
from .text_cleaner import CleanOptions, TextCleaner

DEFAULT_PROFILES_DIR = Path.home() / ".dissertation_docx" / "profiles"


class AppController:
    def __init__(self, profiles_dir: str | Path | None = None):
        self.profiles_dir = Path(profiles_dir) if profiles_dir else DEFAULT_PROFILES_DIR
        self.builder: DocumentBuilder | None = None
        self.dirty = False

    # ============================================================= профили
    def available_profiles(self) -> list[tuple[str, str]]:
        """[(ключ, подпись)]: встроенные пресеты + сохранённые профили-образцы."""
        out = [(f"preset:{key}", factory().name) for key, factory in PRESETS.items()]
        if self.profiles_dir.exists():
            for f in sorted(self.profiles_dir.glob("*.json")):
                try:
                    out.append((f"file:{f}", FormattingProfile.load(f).name))
                except (ValueError, OSError):
                    continue
        return out

    def get_profile(self, key: str | None) -> FormattingProfile:
        key = key or f"preset:{DEFAULT_PRESET}"
        kind, _, value = key.partition(":")
        if kind == "preset":
            return PRESETS[value]()
        if kind == "file":
            return FormattingProfile.load(value)
        raise KeyError(key)

    def load_template(self, path: str | Path, save: bool = True) -> tuple[FormattingProfile, Path | None]:
        """Извлечь профиль из образца .docx и (опционально) сохранить его в JSON."""
        profile = TemplateParser.from_path(path).extract()
        saved = None
        if save:
            self.profiles_dir.mkdir(parents=True, exist_ok=True)
            slug = re.sub(r"[^\w\-]+", "_", Path(path).stem, flags=re.U).strip("_") or "profile"
            saved = self.profiles_dir / f"{slug}.json"
            profile.save(saved)
        return profile, saved

    # ============================================================ документ
    def new_document(self, profile_key: str | None = None) -> DocumentBuilder:
        self.builder = DocumentBuilder.new(self.get_profile(profile_key))
        self.dirty = True
        return self.builder

    def open_document(self, path: str | Path, profile_key: str | None = None,
                      enforce_styles: bool = False) -> DocumentBuilder:
        profile = self.get_profile(profile_key) if profile_key else None
        self.builder = DocumentBuilder.open(path, profile=profile, enforce_styles=enforce_styles)
        self.dirty = enforce_styles
        return self.builder

    def _require(self) -> DocumentBuilder:
        if self.builder is None:
            raise RuntimeError("Документ не открыт")
        return self.builder

    def outline(self):
        return self._require().outline_for_display()

    def preview(self, block: BlockType, text: str, clean: bool = True,
                options: CleanOptions | None = None) -> list[str]:
        """Как текст будет разбит на абзацы (без изменения документа)."""
        cleaner = TextCleaner(options or CleanOptions())
        block = BlockType(block)
        if not clean:
            return [ln for ln in text.split("\n") if ln.strip()]
        if block == BlockType.TITLE_PAGE:
            return cleaner.lines(text, keep_empty=True)
        if block == BlockType.BIBLIOGRAPHY:
            return cleaner.bibliography_entries(text)
        return cleaner.to_paragraphs(text)

    def insert(self, block: BlockType, text: str = "", heading: str = "",
               position: InsertPosition | None = None, clean: bool = True,
               options: CleanOptions | None = None) -> InsertResult:
        result = self._require().insert_block(block, text=text, heading=heading,
                                              position=position, clean=clean, clean_options=options)
        self.dirty = True
        return result

    def save(self, path: str | Path | None = None) -> Path:
        saved = self._require().save(path)
        self.dirty = False
        return saved

    @property
    def path(self) -> Path | None:
        return self.builder.path if self.builder else None
