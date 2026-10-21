@echo off
rem Add "Word equation to MathType" to the .docx right-click menu (HKCU only)
chcp 65001 >nul
setlocal
set "HERE=%~dp0"
set "PY=C:\Users\Yrt\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" "%HERE%install_menu.py" install
echo.
pause
endlocal
