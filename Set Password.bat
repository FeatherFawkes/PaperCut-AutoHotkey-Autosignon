@echo off
setlocal
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File ".\Save-PaperCutCredential.ps1"
echo.
pause
