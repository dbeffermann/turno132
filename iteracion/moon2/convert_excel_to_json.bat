@echo off
setlocal

REM Usage:
REM   convert_excel_to_json.bat [excel_path]
REM Default Excel path:
REM   iteracion\moon2\turno_132_authoring.xlsx

cd /d "%~dp0\..\.."

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
) else (
  set "PY=python"
)

set "XLSX=%~1"
if "%XLSX%"=="" set "XLSX=iteracion\moon2\turno_132_authoring.xlsx"

echo [MVP] Convirtiendo Excel a JSON...
echo [MVP] Excel: %XLSX%
echo [MVP] JSON : iteracion\moon2\game-script.json

"%PY%" tools\narrative_excel_bridge_v2.py import --xlsx "%XLSX%" --json-out iteracion\moon2\game-script.json --json-base iteracion\moon2\game-script.json
if errorlevel 1 (
  echo [ERROR] Fallo la conversion Excel -> JSON.
  exit /b 1
)

echo [OK] Conversion completada.
exit /b 0
