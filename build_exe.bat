@echo off
title Build CRM Pro.exe
cd /d "%~dp0"
setlocal

echo ============================================================
echo   Building CRM Pro.exe (standalone)
echo   This bundles the app + templates + styles into one .exe.
echo ============================================================
echo.

rem ---- Check Python ----
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Install from python.org (tick "Add to PATH").
    pause & exit /b 1
)

rem ---- Install PyInstaller if missing ----
echo Installing PyInstaller (first time only)...
python -m pip install pyinstaller >nul 2>nul

echo.
echo Building one-file .exe (this can take a minute)...
python -m PyInstaller --noconfirm --onefile --name "CRM Pro" ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --hidden-import openpyxl ^
  app.py
if errorlevel 1 goto err

rem ---- Copy the Excel database next to the exe so it persists ----
echo.
echo Copying database next to the exe...
if not exist "dist\data" mkdir "dist\data"
xcopy /E /I /Y "data\*" "dist\data\" >nul

echo.
echo ============================================================
echo   DONE!
echo   Your app is at:  dist\CRM Pro.exe
echo   The data folder (your Excel database) sits next to it.
echo   Double-click "CRM Pro.exe" to run.
echo   NOTE: the .exe works only on Windows (built here).
echo ============================================================
pause
exit /b 0

:err
echo.
echo [ERROR] Build failed. Check the messages above.
pause
exit /b 1
