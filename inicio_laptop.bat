@echo off
cd /d "%~dp0"

echo ==== INICIO %DATE% %TIME% ==== >> update.log

git pull origin main >> update.log 2>&1

start "" pythonw "registro_laptop.pyw"

