@echo off
setlocal

set "SCRIPT_DIR=%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Publish-CorporateSharePoint.ps1" -DryRun -ShowDetails

echo.
echo Preview finished. Review the output above before publishing.
pause

