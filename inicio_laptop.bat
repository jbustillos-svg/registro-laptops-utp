@echo off

REM === MOVERSE A LA CARPETA DONDE ESTA EL .BAT ===
cd /d "%~dp0"

REM === ACTUALIZAR DESDE GITHUB ===
git pull --quiet

REM === RUTA DE PYTHONW ===
set PYTHONW="C:\Users\javie\AppData\Local\Programs\Python\Python311\pythonw.exe"

REM === EJECUTAR SISTEMA ===
start "" %PYTHONW% "registro_laptop.pyw"

exit
