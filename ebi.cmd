@echo off
rem ebi.cmd -- double-click entry point for the ebi day menu.
rem Remembers the last work folder (.ebi\menu.json); pass one to change it:
rem     ebi.cmd C:\work\today
setlocal
set "EBI_DIR=%~dp0"
if "%~1"=="" (
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%EBI_DIR%ebi.ps1" menu
) else (
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%EBI_DIR%ebi.ps1" menu -WorkDir "%~1"
)
if errorlevel 1 pause
endlocal
