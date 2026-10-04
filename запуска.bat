@echo off
setlocal
cd /d "%~dp0"
title Frame House Cost Calculator

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY where py >nul 2>nul && set "PY=py"
if not defined PY (
    echo [ERROR] Python not found.
    echo Install Python from https://www.python.org/downloads/
    echo and check "Add python.exe to PATH" during installation.
    pause
    exit /b 1
)

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing openpyxl...
    %PY% -m pip install openpyxl
    if errorlevel 1 (
        echo [ERROR] Cannot install openpyxl.
        pause
        exit /b 1
    )
)

%PY% karkas_calculator.py
if errorlevel 1 (
    echo.
    echo App closed with error. See details above.
    pause
)
endlocal