@echo off
cd /d "%~dp0"

set "INICIO_BATCH=%DATE% %TIME%"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "bootstrap_laptop.pyw"
) else (
    where pythonw >nul 2>&1
    if not errorlevel 1 (
        start "" pythonw "bootstrap_laptop.pyw"
    ) else (
        where pyw >nul 2>&1
        if not errorlevel 1 (
            start "" pyw -3 "bootstrap_laptop.pyw"
        ) else (
            >>"bootstrap.log" echo [%DATE% %TIME%] No se encontró Python para preparar el sistema.
        )
    )
)
>>"update.log" echo [%DATE% %TIME%] [ARRANQUE] bootstrap solicitado. Inicio del bat: %INICIO_BATCH%
