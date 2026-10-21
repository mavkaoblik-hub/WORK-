@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALL.ps1"
echo.
if errorlevel 1 (
 echo Installer did not complete. Your original XLSX was not overwritten.
)
pause
