@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\pytest.exe" (
  call install.bat dev nopause
  if errorlevel 1 exit /b 1
)
".venv\Scripts\python.exe" -m pytest %*
pause
