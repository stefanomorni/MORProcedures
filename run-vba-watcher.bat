@echo off
set "PY_EXEC=%OFFICE_PYTHON_ENV%"
if "%PY_EXEC%"=="" set "PY_EXEC=C:\Users\Stefano\miniforge3\envs\xlpy\python.exe"

if "%1"=="" (start /min "" "%~f0" run & exit)
:run
title VBA-Watcher — MORProcedures-dev
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
