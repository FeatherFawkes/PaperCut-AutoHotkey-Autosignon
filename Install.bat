@echo off
setlocal
cd /d "%~dp0"
start "" powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File ".\Install-PaperCutAutoLogin.ps1"
exit /b
