@echo off
cd /d "%~dp0"

set "INICIO_BATCH=%DATE% %TIME%"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "registro_laptop.pyw"
) else (
    start "" pythonw "registro_laptop.pyw"
)
>>"update.log" echo [%DATE% %TIME%] [ARRANQUE] pythonw solicitado inmediatamente. Inicio del bat: %INICIO_BATCH%
