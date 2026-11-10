@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem Thin wrapper. Build logic lives in tools/build_release.py because this
rem file must stay pure ASCII (cmd parses .bat in the OEM codepage).
rem Usage: build_exe.bat [--with-resources]

".venv\Scripts\python.exe" "tools\build_release.py" %*
echo.
pause
exit /b 0
