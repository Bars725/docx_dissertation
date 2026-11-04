# docx_dissertation

Настольная программа на Python для оформления диссертации в Word (.docx) по ГОСТ.
Сырой текст вставляется в отдельное окно без инструментов форматирования, а программа
сама применяет оформление нужного раздела: титульного листа, введения, главы,
подглавы, списка литературы или приложения.

* Пресеты: **ГОСТ 7.32-2017** (по умолчанию) и **ГОСТ Р 7.0.11-2011** (диссертация).
* Можно загрузить **образец .docx**: программа извлечёт поля, шрифты, интервалы,
  стили заголовков, формат нумерации, список литературы и колонтитулы и сохранит их как профиль.
* Всё оформление хранится в стилях Word, поэтому документ не «разваливается» при правке
  в Word или LibreOffice, а оглавление и навигация работают.
* Номера глав, подглав, источников и буквы приложений пересчитываются автоматически.

## Установка и запуск (Windows)

1. Python 3.10+ с python.org — с галочками «Add python.exe to PATH» и «tcl/tk and IDLE».
2. `install.bat` — создаст окружение `.venv` и поставит зависимости.
3. `run.bat` — запуск. `run.bat --check` — проверка, всего ли хватает.

Подробности, таблица зависимостей и решение ошибки
`No module named 'exceptions'` (конфликт со старым пакетом `docx`) — в
[docs/INSTALL.md](docs/INSTALL.md).

## Командная строка

```bash
python -m dissertation inspect образец.docx                 # что извлекается из образца
python -m dissertation profile образец.docx -o профиль.json
python -m dissertation new диссертация.docx --preset gost-r-7.0.11-2011
python -m dissertation add диссертация.docx --type chapter --heading "Обзор литературы" --text-file глава1.txt
```

Типы блоков: `title_page`, `introduction`, `chapter`, `subchapter`, `section`, `body`,
`conclusion`, `bibliography`, `appendix`.

## Сборка .exe

```bash
pip install -r requirements-dev.txt
pyinstaller --onefile --windowed --name DissertationDocx --paths src run_app.py
```

## Документация

* [docs/INSTALL.md](docs/INSTALL.md) — установка, зависимости, частые ошибки.
* [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — архитектура, выбор библиотек, рискованные места, план, рекомендации по стабильности python-docx.
* [docs/TESTING.md](docs/TESTING.md) — автотесты, матрица Word / LibreOffice / Google Docs, ручной чек-лист.

## Тесты

```bash
pytest          # тесты с LibreOffice пропускаются, если soffice не установлен
```
