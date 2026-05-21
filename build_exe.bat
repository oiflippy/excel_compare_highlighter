@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "APP_NAME=ExcelCompareHighlighter"
set "BUILD_VENV=.build_venv"
set "PYTHON_CMD=py -3"

echo [1/5] Checking Python...
%PYTHON_CMD% --version >nul 2>nul
if errorlevel 1 (
    set "PYTHON_CMD=python"
    python --version >nul 2>nul
    if errorlevel 1 (
        echo ERROR: Python 3.10 or newer is required to build the EXE.
        pause
        exit /b 1
    )
)

echo [2/5] Creating isolated build environment...
if not exist "%BUILD_VENV%\Scripts\python.exe" (
    %PYTHON_CMD% -m venv "%BUILD_VENV%"
    if errorlevel 1 (
        echo ERROR: Failed to create build virtual environment.
        pause
        exit /b 1
    )
)

echo [3/5] Installing build dependencies...
"%BUILD_VENV%\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :build_failed
"%BUILD_VENV%\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :build_failed

echo [4/5] Building standalone EXE...
"%BUILD_VENV%\Scripts\python.exe" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --icon "favicon.ico" ^
    --add-data "favicon.ico;." ^
    --name "%APP_NAME%" ^
    app.py
if errorlevel 1 goto :build_failed

echo [5/5] Done.
echo Standalone EXE: dist\%APP_NAME%.exe
echo This EXE includes Python and package dependencies. No extra Python install is needed on the target PC.
pause
exit /b 0

:build_failed
echo ERROR: Build failed. Check the messages above.
pause
exit /b 1
