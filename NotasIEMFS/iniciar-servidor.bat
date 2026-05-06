@echo off
chcp 65001 >nul
cd /d "%~dp0"

set PORT=8080
echo.
echo  Servidor local para Sistema de Notas
echo  Carpeta: %CD%
echo  URL:     http://127.0.0.1:%PORT%/
echo.
echo  Deja esta ventana abierta. Para cerrar el servidor, cierrala o pulsa Ctrl+C
echo.

where py >nul 2>&1
if %ERRORLEVEL%==0 (
  start "" "http://127.0.0.1:%PORT%/index.html"
  py -3 "%~dp0servidor_local.py" %PORT%
  goto :fin
)

where python >nul 2>&1
if %ERRORLEVEL%==0 (
  start "" "http://127.0.0.1:%PORT%/index.html"
  python "%~dp0servidor_local.py" %PORT%
  goto :fin
)

echo No se encontro Python. Instala Python desde https://www.python.org/downloads/
echo O abre index.html directamente (el POST ya usa text/plain para evitar CORS).
pause
:fin
