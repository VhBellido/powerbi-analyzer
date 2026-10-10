@echo off
setlocal
cd /d "%~dp0"
echo.
echo Power BI Analyzer V2 - modo manual
echo Use preferencialmente pelo menu Ferramentas Externas do Power BI.
echo.
set /p SERVER=Servidor (ex: localhost:51234): 
set /p DATABASE=Database/GUID: 
if not exist ".venv\Scripts\python.exe" (
  echo Execute INSTALAR_V2.ps1 primeiro.
  pause
  exit /b 2
)
".venv\Scripts\python.exe" powerbi_analyzer_v2.py --server "%SERVER%" --database "%DATABASE%"
pause
