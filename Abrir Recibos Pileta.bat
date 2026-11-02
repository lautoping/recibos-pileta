@echo off
rem Abre la app de Recibos Pileta sin dejar una ventana negra abierta.
cd /d "%~dp0"

where pyw >NUL 2>&1
if %errorlevel%==0 (
    start "" pyw "reemplazo_cuaderno.py"
    exit /b
)

where pythonw >NUL 2>&1
if %errorlevel%==0 (
    start "" pythonw "reemplazo_cuaderno.py"
    exit /b
)

rem Si no hay pythonw, se abre con consola para poder ver cualquier error.
python "reemplazo_cuaderno.py"
if errorlevel 1 (
    echo.
    echo No se pudo abrir la app. Revisa que Python este instalado.
    pause
)
