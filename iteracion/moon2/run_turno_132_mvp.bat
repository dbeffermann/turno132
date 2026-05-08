@echo off
setlocal

REM Run from repo root regardless of where this script is called
cd /d "%~dp0\..\.."

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
) else (
  set "PY=python"
)

echo [1/2] Importando Excel -> JSON runtime...
"%PY%" tools\narrative_excel_bridge_v2.py import --xlsx iteracion\moon2\turno_132_authoring.xlsx --json-out iteracion\moon2\game-script.json --json-base iteracion\moon2\game-script.json
if errorlevel 1 (
  echo [ERROR] Fallo en import. Revisa el workbook y el bridge.
  exit /b 1
)

echo [2/2] Levantando servidor local en http://127.0.0.1:8000 ...
start "" http://127.0.0.1:8000/iteracion/moon2/moon_signal_game.html
"%PY%" -m http.server 8000
