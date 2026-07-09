@echo off
setlocal

set "SCRIPT_DIR=%~dp0"

echo This will publish the Power-BI repo to the ELCC SharePoint PBI folder.
echo Destination:
echo C:\Users\michael.louie\ESDC EDSC\Federal Secretariat on Early Learning and Child Care - Quants\PBI
echo.
choice /M "Continue"
if errorlevel 2 (
    echo Publish cancelled.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Publish-CorporateSharePoint.ps1"

echo.
echo Publish command finished.
pause

