@echo off
set "PY_EXEC=%OFFICE_PYTHON_ENV%"
if "%PY_EXEC%"=="" set "PY_EXEC=C:\Users\Stefano\miniforge3\envs\xlpy\python.exe"

if "%1"=="" (start "" "%~f0" run & exit)
:run
title Ribbon: Force Export (Dev XLSM -> ribbon\)
cd /d "%~dp0"
echo ============================================================
echo  Force Export: Fluent Ribbon XML ^> ribbon\
echo  Use this after: modifying ribbon in CustomUI Editor
echo ============================================================
echo.
"%PY_EXEC%" scripts\ribbon_sync.py export "MORProcedures-dev.xlsm" --output-dir ribbon
echo.
echo === Done. Press any key to close ===
pause
