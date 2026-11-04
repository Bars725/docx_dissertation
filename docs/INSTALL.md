# Установка

## Быстро (Windows)

1. **Python 3.10+** с https://www.python.org/downloads/. В установщике отметьте
   **«Add python.exe to PATH»** и в «Customize installation» — **«tcl/tk and IDLE»**
   (это tkinter, без него не будет окна программы).
2. Двойной щелчок по **`install.bat`**. Он создаст папку `.venv` (отдельное окружение
   только для этой программы) и поставит туда зависимости.
3. Запуск: **`run.bat`**.

Проверить, всё ли на месте: `run.bat --check`.

## Что устанавливается

| Компонент | Зачем | Обязателен | Как ставится |
|---|---|---|---|
| Python ≥ 3.10 | среда выполнения | да | python.org |
| tkinter (tcl/tk) | окно программы | для GUI | вместе с Python (Windows/macOS), `sudo apt install python3-tk` (Linux) |
| **python-docx** ≥ 1.1 | работа с .docx | да | `requirements.txt` |
| lxml | разбор XML документа | да | `requirements.txt` (ставится с python-docx) |
| pytest | автотесты | нет | `requirements-dev.txt` / `install.bat dev` |
| PyInstaller | сборка .exe | нет | `requirements-dev.txt` |
| LibreOffice | тесты совместимости и рендер в PDF | нет | https://www.libreoffice.org/download/ |

`.bat`-файлы:

| Файл | Что делает |
|---|---|
| `install.bat` | `.venv` + зависимости для работы |
| `install.bat dev` | плюс pytest и PyInstaller |
| `run.bat` | запуск программы (при первом запуске сам вызовет установку); `run.bat --check`, `run.bat inspect образец.docx` |
| `test.bat` | автотесты |
| `push.bat` | применить пакет обновления из `_claude_update`, при необходимости закоммитить правки и отправить на GitHub (`git push`) с подтверждением |

## Ошибка `ModuleNotFoundError: No module named 'exceptions'`

В Python установлен пакет **`docx`** (2014 г., только для Python 2). Он занимает то же имя
импорта `docx`, что и нужный **`python-docx`**. Варианты:

* **рекомендуется**: запускать через `install.bat` / `run.bat`. В `.venv` конфликтного
  пакета нет, а глобальный Python остаётся нетронутым;
* или починить глобальный Python:

  ```bat
  python -m pip uninstall -y docx
  python -m pip install --force-reinstall python-docx
  ```

## Вручную (любая ОС)

```bash
python -m venv .venv
.venv\Scripts\activate                # Linux/macOS: source .venv/bin/activate
python -m pip install -r requirements.txt        # или requirements-dev.txt
python run_app.py --check
python run_app.py
```
