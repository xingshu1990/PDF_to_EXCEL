@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem ---------------------------------------------------------------
rem  BATCH MAPPING WORKBENCH: map every field of a PDF to a template
rem  column in one table (no more dragging cell by cell)
rem  Usage: drag a PDF onto this file, or double-click and press Enter
rem ---------------------------------------------------------------
set "PY=%~dp0.venv\Scripts\python.exe"
if exist "%PY%" goto run
set "PY=py"
where py >nul 2>nul || set "PY=python"
:run
"%PY%" "%~dp0pdf_table_master.py" --map %*
echo.
pause
