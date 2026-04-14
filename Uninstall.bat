@echo off
setlocal
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File ".\Uninstall-PaperCutAutoLogin.ps1"
echo.
pause
