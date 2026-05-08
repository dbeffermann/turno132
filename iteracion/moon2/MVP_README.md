# Turno 132 MVP

Este MVP queda reducido a los archivos operativos para editar, convertir y jugar.

## Archivos finales del MVP

- Excel de autoría: `iteracion/moon2/turno_132_authoring.xlsx`
- Runtime UI: `iteracion/moon2/moon_signal_game.html`
- Script runtime (generado desde Excel): `iteracion/moon2/game-script.json`
- Bridge de conversión: `tools/narrative_excel_bridge_v2.py`
- Assets del juego: `iteracion/moon2/assets/`
- Lanzador MVP (import + server + abrir juego): `iteracion/moon2/run_turno_132_mvp.bat`

## Flujo recomendado (único)

1. Editar el contenido en `turno_132_authoring.xlsx`.
2. Ejecutar `run_turno_132_mvp.bat`.
3. El script:
   - importa Excel -> `game-script.json`
   - abre `http://127.0.0.1:8000/iteracion/moon2/moon_signal_game.html`
   - deja el servidor local activo.

## Notas técnicas

- El SFX de notificación apunta a un asset existente:
  - `iteracion/moon2/assets/audio/sfx/inicio docu.mp3`
- Si usas entorno virtual, el .bat prioriza `.venv/Scripts/python.exe` automáticamente.
