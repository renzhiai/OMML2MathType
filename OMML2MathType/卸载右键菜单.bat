@echo off
rem Remove "Word equation to MathType" from the .docx right-click menu
chcp 65001 >nul
setlocal
set "HERE=%~dp0"
set "PY=C:\Users\Yrt\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" "%HERE%install_menu.py" uninstall
echo.
pause
endlocal
