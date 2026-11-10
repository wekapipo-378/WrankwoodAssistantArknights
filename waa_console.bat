@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

".venv\Scripts\python.exe" "tools\axis_console.py" %*
rem exit code 2 = another console is already running and was brought to front.
rem Only that exact code skips the prompt; crashes still show the traceback.
if "%ERRORLEVEL%"=="2" exit /b 0
pause
