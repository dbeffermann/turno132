@echo off
setlocal
cd /d "%~dp0"

echo [1/3] Verificando remoto origin...
git remote get-url origin >nul 2>&1
if errorlevel 1 (
  echo [ERROR] No existe remoto 'origin'.
  echo Configuralo primero, por ejemplo:
  echo   git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
  exit /b 1
)

echo [2/3] Publicando branch master...
git push -u origin master
if errorlevel 1 (
  echo [ERROR] Fallo el push. Revisa permisos/autenticacion.
  exit /b 1
)

echo [3/3] Listo. El workflow de Pages quedó en .github/workflows/pages.yml
echo Si es primera vez, en GitHub activa: Settings -^> Pages -^> Source: GitHub Actions
echo URL esperada: https://TU_USUARIO.github.io/TU_REPO/iteracion/moon2/moon_signal_game.html

exit /b 0
