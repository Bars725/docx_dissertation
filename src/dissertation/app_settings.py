"""Настройки пользователя: последняя папка с изображениями и закреплённые за файлами подписи.

Хранятся в ``~/.dissertation_docx/settings.json`` (рядом с профилями), а не в папке
с картинками — чтобы не мусорить в ней. Ключ подписи — полный путь к файлу.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

from .figures import IMAGE_EXT, clean_title


class AppSettings:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.data: dict = {}
        try:
            self.data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            self.data = {}

    def save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    # --------------------------------------------------------------- папка
    @property
    def image_folder(self) -> str:
        return self.data.get("image_folder", "")

    @image_folder.setter
    def image_folder(self, value: str) -> None:
        self.data["image_folder"] = value
        self.save()

    # -------------------------------------------------------------- подписи
    @staticmethod
    def _key(path: str) -> str:
        return os.path.normcase(os.path.abspath(path))

    def caption(self, path: str) -> str:
        return self.data.get("captions", {}).get(self._key(path), "")

    def set_caption(self, path: str, title: str) -> None:
        caps = self.data.setdefault("captions", {})
        title = clean_title(title)
        if title:
            caps[self._key(path)] = title
        else:
            caps.pop(self._key(path), None)
        self.save()


def list_images(folder: str) -> list[str]:
    """Изображения в папке (без вложенных), по имени без учёта регистра."""
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    files = [os.path.join(folder, n) for n in names
             if n.lower().endswith(IMAGE_EXT) and os.path.isfile(os.path.join(folder, n))]
    return sorted(files, key=lambda p: os.path.basename(p).casefold())


def default_title(path: str) -> str:
    """Подпись по умолчанию из имени файла: «schema_bd-v2.png» → «Schema bd v2»."""
    stem = os.path.splitext(os.path.basename(path))[0]
    return clean_title(stem.replace("_", " ").replace("-", " "))
