@echo off

REM === MOVERSE A LA CARPETA DEL SISTEMA ===
cd /d "%~dp0"

REM === ACTUALIZAR DESDE GITHUB (SI HAY INTERNET) ===
git pull --quiet

REM === BUSCAR PYTHONW AUTOMATICAMENTE ===
where pythonw >nul 2>&1
if errorlevel 1 (
    echo Python no encontrado.
    pause
    exit
)

REM === EJECUTAR SISTEMA ===
start "" pythonw "registro_laptop.pyw"

exit
