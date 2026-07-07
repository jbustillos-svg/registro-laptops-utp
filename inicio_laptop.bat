@echo off
cd /d "%~dp0"

start "" pythonw "registro_laptop.pyw"

start "" /b cmd /c "echo ==== INICIO %DATE% %TIME% ==== >> update.log && git pull origin main >> update.log 2>&1"