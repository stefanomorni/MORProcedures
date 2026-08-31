@echo off
set "PY_EXEC=%OFFICE_PYTHON_ENV%"
if "%PY_EXEC%"=="" set "PY_EXEC=C:\Users\Stefano\miniforge3\envs\xlpy\python.exe"

if "%1"=="" (start "" "%~f0" run & exit)
:run
title PQ: Force Export (Dev XLSM -> power-queries)
cd /d "%~dp0"
echo ============================================================
echo  Force Export: Queries ^> power-queries\
echo  (Add-ins do not use PQ — this launcher is provisioned for
echo   uniformity but will find no queries to export)
echo ============================================================
echo.
"%PY_EXEC%" scripts\pq_sync.py export-all "MORProcedures-dev.xlsm"
echo.
echo === Done. Press any key to close ===
pause
