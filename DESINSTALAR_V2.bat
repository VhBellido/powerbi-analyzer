@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0DESINSTALAR_FERRAMENTA_EXTERNA.ps1"
pause
