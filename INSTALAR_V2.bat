@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALAR_V2.ps1"
if errorlevel 1 (
  echo.
  echo A instalacao terminou com erro.
  pause
)
