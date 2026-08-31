@echo off
set "PY_EXEC=%OFFICE_PYTHON_ENV%"
if "%PY_EXEC%"=="" set "PY_EXEC=C:\Users\Stefano\miniforge3\envs\xlpy\python.exe"

if "%1"=="" (start /min "" "%~f0" run & exit)
:run
title PQ-Watcher — MORProcedures-dev
cd /d "%~dp0"
echo ============================================================
echo  PQ Watcher: MORProcedures-dev.xlsm ^<^> power-queries\
echo  (Add-ins do not use PQ — watcher will run but find nothing)
echo ============================================================
echo.
"%PY_EXEC%" scripts\pq_sync_watch.py watch "MORProcedures-dev.xlsm" --pq-dir power-queries --sync-script scripts\pq_sync.py
echo.
echo === Watcher stopped. Press any key to close ===
pause
