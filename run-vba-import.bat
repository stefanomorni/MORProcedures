@echo off
set "PY_EXEC=%OFFICE_PYTHON_ENV%"
if "%PY_EXEC%"=="" set "PY_EXEC=C:\Users\Stefano\miniforge3\envs\xlpy\python.exe"

if "%1"=="" (start "" "%~f0" run & exit)
:run
title VBA: Force Import (vba-files -> Dev XLSM)
cd /d "%~dp0"
echo ============================================================
echo  Force Import: vba-files\ ^> MORProcedures-dev VBProject
echo  NOTE: Target must be MORProcedures-dev.xlsm, NOT the .xlam
echo ============================================================
echo.
"%PY_EXEC%" scripts\vba_sync_watch.py import-all "MORProcedures-dev.xlsm"
echo.
echo === Done. Press any key to close ===
pause
