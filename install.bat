@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
rem ================================================================
rem  Установка: отдельное окружение .venv + все зависимости.
rem  install.bat       - для работы программы
rem  install.bat dev   - плюс pytest и PyInstaller
rem ================================================================

set "PY="
where py >nul 2>nul && set "PY=py -3"
if defined PY goto :have_py
where python >nul 2>nul && set "PY=python"
if defined PY goto :have_py
echo Python не найден.
echo Установите Python 3.10+ с https://www.python.org/downloads/
echo При установке отметьте "Add python.exe to PATH" и "tcl/tk and IDLE".
goto :fail

:have_py
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"
if errorlevel 1 (
  echo Нужен Python 3.10 или новее.
  goto :fail
)

if exist ".venv\Scripts\python.exe" goto :have_venv
echo [1/4] Создаю виртуальное окружение .venv ...
%PY% -m venv .venv
if errorlevel 1 goto :fail

:have_venv
set "VPY=.venv\Scripts\python.exe"
echo [2/4] Обновляю pip ...
"%VPY%" -m pip install --upgrade pip --disable-pip-version-check -q
if errorlevel 1 goto :fail

rem Устаревший пакет docx для Python 2 конфликтует с python-docx
"%VPY%" -m pip uninstall -y docx >nul 2>nul

echo [3/4] Устанавливаю зависимости ...
"%VPY%" -m pip install -r requirements.txt --disable-pip-version-check
if errorlevel 1 goto :fail
if /i "%~1"=="dev" "%VPY%" -m pip install -r requirements-dev.txt --disable-pip-version-check
if errorlevel 1 goto :fail

echo [4/4] Проверяю окружение ...
"%VPY%" run_app.py --check
if errorlevel 1 goto :fail

echo.
echo Готово. Запуск программы: run.bat
if /i not "%~2"=="nopause" pause
exit /b 0

:fail
echo.
echo Установка не завершена - см. сообщения выше.
pause
exit /b 1
