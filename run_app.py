"""Запуск без установки пакета и точка входа для PyInstaller.

    python run_app.py
    pyinstaller --onefile --windowed --name DissertationDocx --paths src run_app.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from dissertation.__main__ import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
