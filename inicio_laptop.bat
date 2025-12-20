@echo off

REM === RUTA DEL SISTEMA ===
cd /d C:\ProgramData\UTP\RegistroLaptop

REM === ACTUALIZAR DESDE GITHUB ===
git pull --quiet

REM === EJECUTAR SISTEMA ===
set PYTHONW="C:\Users\javie\AppData\Local\Programs\Python\Python311\pythonw.exe"
start "" %PYTHONW% "C:\ProgramData\UTP\RegistroLaptop\registro_laptop.pyw"

exit
