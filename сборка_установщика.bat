@echo off
setlocal
cd /d "%~dp0"
title Build Installer

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY where py >nul 2>nul && set "PY=py"
if not defined PY (
    echo [ERROR] Python not found.
    pause
    exit /b 1
)

%PY% -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    %PY% -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] Cannot install PyInstaller.
        pause
        exit /b 1
    )
)

%PY% -m PyInstaller --noconfirm --clean --onedir --windowed --icon=icon.ico --name KarkasCalculator karkas_calculator.py
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    pause
    exit /b 1
)

set "ISCC="
for %%P in (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
    "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
) do if not defined ISCC if exist %%P set "ISCC=%%~P"

if not defined ISCC (
    echo [ERROR] Inno Setup 6 not found. Install it from https://jrsoftware.org/isinfo.php
    pause
    exit /b 1
)

"%ISCC%" setup.iss
if errorlevel 1 (
    echo [ERROR] Installer build failed.
    pause
    exit /b 1
)

echo.
echo Done. Installer: installer\Калькулятор_каркасного_дома_Установщик.exe
echo Portable build:  dist\KarkasCalculator\KarkasCalculator.exe
pause
endlocal
