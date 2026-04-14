@echo off
setlocal
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File ".\Install-PaperCutAutoLogin.ps1"
echo.
pause
