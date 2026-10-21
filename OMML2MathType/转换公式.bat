@echo off
rem ============================================================
rem  OMML2MathType - drag and drop your .docx here
rem  Converts Word built-in equations (OMML) into MathType objects
rem ============================================================
chcp 65001 >nul
setlocal enabledelayedexpansion
set "HERE=%~dp0"
set "PY=C:\Users\Yrt\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" -c "import docx, win32com.client" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [!] Missing Python packages. Run:
    echo     "%PY%" -m pip install python-docx pywin32 lxml
    echo.
    pause
    exit /b 1
)

if "%~1"=="" goto ask

:loop
if "%~1"=="" goto done
echo.
echo ============================================================
echo   File: %~nx1
echo ============================================================
"%PY%" "%HERE%omml2mathtype.py" "%~1" --inplace --report "%HERE%_report.json"
shift
goto loop

:ask
echo.
set /p "F=Drag a .docx file into this window, then press Enter: "
if "%F%"=="" goto done
set "F=%F:"=%"
"%PY%" "%HERE%omml2mathtype.py" "%F%" --inplace --report "%HERE%_report.json"

:done
echo.
echo All done. Press any key to close.
pause >nul
endlocal
