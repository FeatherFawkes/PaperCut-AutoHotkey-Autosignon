@echo off
setlocal
cd /d "%~dp0"
start "" powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File ".\Uninstall-PaperCutAutoLogin.ps1"
exit /b
