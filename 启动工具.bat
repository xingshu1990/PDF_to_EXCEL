@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem ---------------------------------------------------------------
rem  ONE-CLICK MODE: PDF -> quotation sheet
rem  Usage: drag a PDF onto this file, or double-click and press Enter
rem  For all features (batch / review window / audit ...) use the menu bat
rem ---------------------------------------------------------------
set "PY=%~dp0.venv\Scripts\python.exe"
if exist "%PY%" goto run
set "PY=py"
where py >nul 2>nul || set "PY=python"
:run
"%PY%" "%~dp0pdf_table_master.py" %*
echo.
pause
