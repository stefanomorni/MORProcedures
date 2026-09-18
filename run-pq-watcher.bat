@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start /min "" "%~f0" run & exit)
:run
title PQ-Watcher ? MORProcedures-dev
cd /d "%~dp0"
echo ============================================================
echo  PQ Watcher: MORProcedures-dev.xlsm ^<^> power-queries\
echo  (Add-ins do not use PQ ? watcher will run but find nothing)
echo ============================================================
echo.
"%PY_EXEC%" scripts\pq_sync_watch.py watch "MORProcedures-dev.xlsm" --pq-dir power-queries --sync-script scripts\pq_sync.py
echo.
echo === Watcher stopped. Press any key to close ===
pause
