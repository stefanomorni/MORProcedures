@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start /min "" "%~f0" run & exit)
:run
title VBA-Watcher - MORProcedures-dev
cd /d "%~dp0"
echo ============================================================
echo  VBA Watcher: MORProcedures-dev.xlsm ^<^> vba-files\
echo  NOTE: Always target MORProcedures-dev.xlsm, NOT the .xlam
echo  When ready to publish: run run-addin-publish.bat
echo ============================================================
echo.
"%PY_EXEC%" scripts\vba_sync_watch.py watch "MORProcedures-dev.xlsm" --vba-dir vba-files
echo.
echo === Watcher stopped. Press any key to close ===
pause
