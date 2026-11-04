"""Запуск без установки пакета и точка входа для PyInstaller.

    python run_app.py            — GUI
    python run_app.py --check    — проверить окружение (что установить)
    python run_app.py inspect образец.docx   — команды CLI, см. README
    pyinstaller --onefile --windowed --name DissertationDocx --paths src run_app.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from dissertation.envcheck import check_environment, format_report, has_errors  # noqa: E402


def _show_startup_error(text: str) -> None:
    print(text, file=sys.stderr)
    try:  # при запуске двойным щелчком консоль может закрыться — дублируем окном
        import tkinter
        from tkinter import messagebox

        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Диссертация по ГОСТ", text)
        root.destroy()
    except Exception:  # noqa: BLE001
        pass


def start() -> int:
    args = sys.argv[1:]
    if args[:1] == ["--check"]:
        problems = check_environment(need_gui=True)
        print(format_report(problems))
        return 1 if has_errors(problems) else 0
    problems = check_environment(need_gui=not args)
    if has_errors(problems):
        _show_startup_error(format_report([p for p in problems if p.level != "info"]))
        return 1
    from dissertation.__main__ import main

    return main(args)


if __name__ == "__main__":
    raise SystemExit(start())
