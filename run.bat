@echo off
title CRM Pro - Auto Launcher
cd /d "%~dp0"
setlocal enabledelayedexpansion

echo ============================================================
echo   CRM Pro  -  one-click setup and launch
echo ============================================================

rem ---- Check Python is installed ----
where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo  [ERROR] Python was not found.
    echo  Please install Python from https://www.python.org/downloads/
    echo  and tick "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

rem ---- Auto-install dependencies if missing ----
echo.
echo  Checking dependencies...
python -c "import flask, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo  First run - installing requirements (flask, openpyxl)...
    python -m pip install --upgrade pip >nul 2>nul
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo  [ERROR] Failed to install dependencies.
        echo  Check your internet connection and try again.
        pause
        exit /b 1
    )
) else (
    echo  Dependencies OK.
)

echo.
echo  Starting CRM Pro...
echo  Your browser will open automatically at: http://127.0.0.1:5000
echo  Close this window (or press Ctrl+C) to stop the server.
echo ============================================================

python app.py
pause
