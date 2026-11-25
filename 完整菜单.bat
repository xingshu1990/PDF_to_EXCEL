@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem ---------------------------------------------------------------
rem  FULL MENU (11 items): fill / batch / review window / audit ...
rem  If you only need the quotation sheet, use the one-click bat
rem ---------------------------------------------------------------
set "PY=%~dp0.venv\Scripts\python.exe"
if exist "%PY%" goto run
set "PY=py"
where py >nul 2>nul || set "PY=python"
:run
"%PY%" "%~dp0pdf_table_master.py" --menu
echo.
pause
