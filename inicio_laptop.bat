@echo off
cd /d "%~dp0"

set "INICIO_BATCH=%DATE% %TIME%"
>>"arranque.log" echo [%DATE% %TIME%] BAT iniciado ^| cwd=%CD%

set "PYTHON_BOOTSTRAP="
set "PYTHON_ARGUMENTOS="

if exist "%~dp0.venv\Scripts\pythonw.exe" set "PYTHON_BOOTSTRAP=%~dp0.venv\Scripts\pythonw.exe"

if not defined PYTHON_BOOTSTRAP (
    for /f "delims=" %%P in ('where pythonw 2^>nul') do if not defined PYTHON_BOOTSTRAP set "PYTHON_BOOTSTRAP=%%P"
)

if not defined PYTHON_BOOTSTRAP (
    for /f "delims=" %%P in ('where pyw 2^>nul') do if not defined PYTHON_BOOTSTRAP (
        set "PYTHON_BOOTSTRAP=%%P"
        set "PYTHON_ARGUMENTOS=-3"
    )
)

if not defined PYTHON_BOOTSTRAP (
    for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do if not defined PYTHON_BOOTSTRAP if exist "%%~fD\pythonw.exe" set "PYTHON_BOOTSTRAP=%%~fD\pythonw.exe"
)

if not defined PYTHON_BOOTSTRAP (
    for /d %%D in ("%ProgramFiles%\Python*") do if not defined PYTHON_BOOTSTRAP if exist "%%~fD\pythonw.exe" set "PYTHON_BOOTSTRAP=%%~fD\pythonw.exe"
)

if not defined PYTHON_BOOTSTRAP (
    >>"arranque.log" echo [%DATE% %TIME%] ERROR: Python no encontrado
    exit /b 1
)

>>"arranque.log" echo [%DATE% %TIME%] Python encontrado ^| %PYTHON_BOOTSTRAP% %PYTHON_ARGUMENTOS%
>>"arranque.log" echo [%DATE% %TIME%] bootstrap solicitado
start "" "%PYTHON_BOOTSTRAP%" %PYTHON_ARGUMENTOS% "%~dp0bootstrap_laptop.pyw"
if errorlevel 1 >>"arranque.log" echo [%DATE% %TIME%] ERROR al solicitar bootstrap ^| codigo=%ERRORLEVEL%
>>"update.log" echo [%DATE% %TIME%] [ARRANQUE] bootstrap solicitado. Inicio del bat: %INICIO_BATCH%
