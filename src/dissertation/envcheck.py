"""Проверка окружения до импорта python-docx.

Самая частая поломка — пакет ``docx`` (PyPI, 2014 г., только Python 2), который
занимает то же имя импорта, что и ``python-docx``, и падает с
``ModuleNotFoundError: No module named 'exceptions'``. Модуль не импортирует
docx вслепую, а сначала смотрит, *какой* файл будет загружен.
"""
from __future__ import annotations

import importlib
import importlib.util
import shutil
import sys
from dataclasses import dataclass

MIN_PYTHON = (3, 10)
MIN_PYTHON_DOCX = (1, 1)


@dataclass
class Problem:
    level: str      # error | warning | info
    title: str
    fix: str = ""


def _pip(args: str) -> str:
    return f'"{sys.executable}" -m pip {args}'


def _version_tuple(text: str) -> tuple[int, ...]:
    out = []
    for part in text.split("."):
        digits = ""
        for ch in part:
            if not ch.isdigit():
                break
            digits += ch
        if not digits:
            break
        out.append(int(digits))
        if len(digits) < len(part):   # «0rc1» — дальше суффикс пререлиза
            break
    return tuple(out)


def check_docx() -> list[Problem]:
    spec = importlib.util.find_spec("docx")
    if spec is None:
        return [Problem("error", "Не установлен python-docx", _pip("install -r requirements.txt"))]
    origin = spec.origin or ""
    # python-docx — пакет (docx/__init__.py); конфликтующий «docx» — одиночный файл docx.py.
    if spec.submodule_search_locations is None or origin.lower().endswith("docx.py"):
        return [Problem(
            "error",
            f"Установлен устаревший пакет «docx» (Python 2), он мешает python-docx:\n    {origin}",
            f"{_pip('uninstall -y docx')}\n{_pip('install --force-reinstall python-docx')}\n"
            "или запустите install.bat — он создаст отдельное окружение .venv без конфликта")]
    try:
        docx = importlib.import_module("docx")
    except Exception as exc:  # noqa: BLE001
        return [Problem("error", f"python-docx не импортируется: {type(exc).__name__}: {exc}",
                        _pip("install --force-reinstall python-docx"))]
    version = getattr(docx, "__version__", "0")
    if not hasattr(docx, "Document") or _version_tuple(version) < MIN_PYTHON_DOCX:
        return [Problem("error", f"Слишком старый python-docx ({version}), нужен ≥ 1.1",
                        _pip("install --upgrade python-docx"))]
    return []


def check_environment(need_gui: bool = True) -> list[Problem]:
    problems: list[Problem] = []
    if sys.version_info < MIN_PYTHON:
        problems.append(Problem(
            "error", f"Python {sys.version.split()[0]} — нужен 3.10 или новее",
            "Установите Python с https://www.python.org/downloads/"))
    problems += check_docx()
    try:
        importlib.import_module("lxml.etree")
    except Exception:  # noqa: BLE001
        problems.append(Problem("error", "Не установлен lxml", _pip("install lxml")))
    try:
        importlib.import_module("tkinter")
    except Exception:  # noqa: BLE001
        fix = ("Windows: переустановите Python, отметив «tcl/tk and IDLE»; "
               "Linux: sudo apt install python3-tk; macOS: brew install python-tk")
        problems.append(Problem("error" if need_gui else "warning",
                                "Нет tkinter — графический интерфейс недоступен", fix))
    if importlib.util.find_spec("PIL") is None:
        problems.append(Problem("warning", "Не установлен Pillow — предпросмотр JPG/BMP/TIFF в окне "
                                           "изображений недоступен (PNG/GIF работают)",
                                _pip("install Pillow")))
    if importlib.util.find_spec("pytest") is None:
        problems.append(Problem("info", "pytest не установлен (нужен только для тестов)",
                                _pip("install -r requirements-dev.txt")))
    if not (shutil.which("soffice") or shutil.which("libreoffice")):
        problems.append(Problem("info", "LibreOffice не найден — тесты совместимости с ним будут пропущены"))
    return problems


def format_report(problems: list[Problem]) -> str:
    if not any(p.level == "error" for p in problems):
        head = f"Окружение в порядке (Python {sys.version.split()[0]}, {sys.executable})"
    else:
        head = "Программа не может запуститься:"
    marks = {"error": "✗", "warning": "!", "info": "·"}
    lines = [head]
    for p in problems:
        lines.append(f"  {marks[p.level]} {p.title}")
        if p.fix:
            lines += ["      " + ln for ln in ("Исправить:\n" + p.fix).split("\n")]
    return "\n".join(lines)


def has_errors(problems: list[Problem]) -> bool:
    return any(p.level == "error" for p in problems)
