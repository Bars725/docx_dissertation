@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
rem ================================================================
rem  Отправка изменений на GitHub (origin).
rem  1. применяет пакет обновления из _claude_update, если он есть;
rem  2. предлагает закоммитить несохранённые изменения;
rem  3. проверяет, нет ли новых коммитов на GitHub;
rem  4. показывает, что будет отправлено, и делает git push.
rem ================================================================

where git >nul 2>nul
if errorlevel 1 goto :no_git

for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD') do set "BR=%%b"
if not defined BR goto :not_repo
echo Ветка: %BR%
echo.

rem ---- 1. пакет обновления
if not exist "_claude_update\docx_dissertation.bundle" goto :check_dirty
git bundle verify -q "_claude_update\docx_dissertation.bundle" >nul 2>nul
if errorlevel 1 goto :check_dirty
echo Применяю пакет обновления _claude_update\docx_dissertation.bundle ...
git pull --ff-only -q "_claude_update\docx_dissertation.bundle" main
if errorlevel 1 goto :bundle_fail

:check_dirty
rem ---- 2. несохранённые изменения
set "DIRTY="
for /f "delims=" %%i in ('git status --porcelain') do set "DIRTY=1"
if not defined DIRTY goto :fetch
echo Есть незакоммиченные изменения:
git status --short
echo.
choice /C YN /M "Закоммитить их перед отправкой"
if errorlevel 2 goto :fetch
set "MSG="
set /p "MSG=Сообщение коммита (Enter - «Обновление»): "
if not defined MSG set "MSG=Обновление"
git add -A
git commit -q -m "%MSG%"
if errorlevel 1 goto :fail
echo.

:fetch
rem ---- 3. что на GitHub
echo Проверяю GitHub ...
git fetch -q origin
if errorlevel 1 goto :auth_fail
git rev-parse --verify -q "origin/%BR%" >nul
if errorlevel 1 goto :first_push

for /f "tokens=1,2" %%a in ('git rev-list --left-right --count "origin/%BR%...%BR%"') do (
  set "BEHIND=%%a"
  set "AHEAD=%%b"
)
if not "%BEHIND%"=="0" goto :behind
if "%AHEAD%"=="0" goto :nothing

rem ---- 4. отправка
echo.
echo Будет отправлено коммитов: %AHEAD%
git log --oneline "origin/%BR%..%BR%"
echo.
choice /C YN /M "Отправить на GitHub"
if errorlevel 2 goto :cancel
git push origin "%BR%"
if errorlevel 1 goto :auth_fail
goto :done

:first_push
echo Ветки %BR% на GitHub ещё нет - будет создана.
git log --oneline -10
choice /C YN /M "Отправить на GitHub"
if errorlevel 2 goto :cancel
git push -u origin "%BR%"
if errorlevel 1 goto :auth_fail
goto :done

:behind
echo.
echo На GitHub есть новые коммиты, которых нет локально: %BEHIND%
echo Пробую подтянуть их без слияния...
git pull --ff-only -q origin "%BR%"
if errorlevel 1 goto :diverged
goto :fetch

:diverged
echo История на GitHub и локально разошлась.
echo Разрешите вручную, например:  git pull --rebase origin %BR%
goto :fail

:nothing
echo Нечего отправлять: GitHub уже содержит все локальные коммиты.
goto :end

:done
echo.
echo Готово: изменения отправлены на GitHub.
goto :end

:cancel
echo Отменено.
goto :end

:bundle_fail
echo Не удалось применить пакет обновления.
echo Возможно, локальные изменения мешают: закоммитьте или отмените их и запустите снова.
goto :fail

:auth_fail
echo.
echo Не удалось связаться с GitHub.
echo  - проверьте интернет;
echo  - при первом push должно открыться окно входа Git Credential Manager;
echo    если его нет - установите Git for Windows заново или выполните: gh auth login
goto :fail

:not_repo
echo Эта папка не является git-репозиторием.
goto :fail

:no_git
echo Git не найден. Установите Git for Windows: https://git-scm.com/download/win
goto :fail

:fail
echo.
echo Отправка не выполнена.
pause
exit /b 1

:end
pause
exit /b 0
