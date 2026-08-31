@echo off
set "PY_EXEC=%OFFICE_PYTHON_ENV%"
if "%PY_EXEC%"=="" set "PY_EXEC=C:\Users\Stefano\miniforge3\envs\xlpy\python.exe"

if "%1"=="" (start "" "%~f0" run & exit)
:run
title Add-In Publish: MORProcedures-dev.xlsm -> MORProcedures.xlam
cd /d "%~dp0"
echo ============================================================
echo  PUBLISH ADD-IN
echo  Source : MORProcedures-dev.xlsm  (dev workbook)
echo  Target : MORProcedures.xlam      (canonical add-in)
echo ============================================================
echo.
echo [1/3] Exporting current VBA state to vba-files\ ...
"%PY_EXEC%" scripts\vba_sync_watch.py export-all "MORProcedures-dev.xlsm"
if errorlevel 1 (
    echo [ERROR] VBA export failed. Aborting publish.
    pause & exit /b 1
)
echo.
echo [2/3] Saving dev workbook as staged add-in ...
"%PY_EXEC%" -c "
import win32com.client as win32, os, shutil, pathlib
xl = win32.GetActiveObject('Excel.Application')
wb = None
for w in xl.Workbooks:
    if w.Name.lower() == 'morprocedures-dev.xlsm':
        wb = w
        break
if wb is None:
    print('ERROR: MORProcedures-dev.xlsm not found in Excel. Open it first.')
    exit(1)
stage = str(pathlib.Path(r'%~dp0') / '_stage_MORProcedures.xlam')
xl.DisplayAlerts = False
wb.SaveAs(stage, FileFormat=55)  # 55 = xlAddIn
xl.DisplayAlerts = True
print(f'Staged: {stage}')
"
if errorlevel 1 (
    echo [ERROR] SaveAs xlAddIn failed. Aborting publish.
    pause & exit /b 1
)
echo.
echo [3/3] Copying staged add-in to canonical location ...
"%PY_EXEC%" -c "
import shutil, pathlib
stage  = pathlib.Path(r'%~dp0_stage_MORProcedures.xlam')
target = pathlib.Path(r'%~dp0MORProcedures.xlam').resolve()
shutil.copy2(str(stage), str(target))
stage.unlink()
print(f'Published: {target}')
"
if errorlevel 1 (
    echo [ERROR] Copy to canonical .xlam failed.
    pause & exit /b 1
)
echo.
echo ============================================================
echo  Publish complete. Reload the add-in in Excel if needed:
echo  File ^> Options ^> Add-Ins ^> Manage: Excel Add-Ins ^> Browse
echo ============================================================
pause
