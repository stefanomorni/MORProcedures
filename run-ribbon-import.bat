@echo off
set "PY_EXEC=%OFFICE_PYTHON_ENV%"
if "%PY_EXEC%"=="" set "PY_EXEC=C:\Users\Stefano\miniforge3\envs\xlpy\python.exe"

if "%1"=="" (start "" "%~f0" run & exit)
:run
title Ribbon: Force Import (ribbon\ -> Dev XLSM)
cd /d "%~dp0"
echo ============================================================
echo  Force Import: ribbon\ ^> Fluent Ribbon XML in Dev XLSM
echo  Use this after: editing customUI14.xml on disk
echo ============================================================
echo.
"%PY_EXEC%" scripts\ribbon_sync.py import "MORProcedures-dev.xlsm" --input-dir ribbon
echo.
echo === Done. Press any key to close ===
pause
