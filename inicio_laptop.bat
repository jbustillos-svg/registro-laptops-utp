@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "INICIO_BATCH=%DATE% %TIME%"
set "RUTA_LOG=%~dp0arranque.log"
set "MAX_INTENTOS=3"
set "ESPERA_REINTENTO=2"
set "RAIZ_LOCALAPPDATA=%LocalAppData%"
set "RAIZ_PROGRAMFILES=%ProgramFiles%"
set "RAIZ_PROGRAMFILES_X86=%ProgramFiles(x86)%"
set "RAIZ_WINDOWS=%SystemRoot%"
if defined REGISTRO_LAPTOP_LOCALAPPDATA set "RAIZ_LOCALAPPDATA=%REGISTRO_LAPTOP_LOCALAPPDATA%"
if defined REGISTRO_LAPTOP_PROGRAMFILES set "RAIZ_PROGRAMFILES=%REGISTRO_LAPTOP_PROGRAMFILES%"
if defined REGISTRO_LAPTOP_PROGRAMFILES_X86 set "RAIZ_PROGRAMFILES_X86=%REGISTRO_LAPTOP_PROGRAMFILES_X86%"
if defined REGISTRO_LAPTOP_WINDOWS set "RAIZ_WINDOWS=%REGISTRO_LAPTOP_WINDOWS%"
if defined REGISTRO_LAPTOP_REINTENTO_SEGUNDOS set "ESPERA_REINTENTO=%REGISTRO_LAPTOP_REINTENTO_SEGUNDOS%"
>>"%RUTA_LOG%" echo [%DATE% %TIME%] BAT iniciado ^| cwd=%CD%

for /l %%I in (1,1,%MAX_INTENTOS%) do (
    set "PYTHON_BOOTSTRAP="
    set "PYTHON_ARGUMENTOS="
    call :buscar_python %%I
    if defined PYTHON_BOOTSTRAP goto :python_encontrado
    if %%I lss %MAX_INTENTOS% (
        >>"%RUTA_LOG%" echo [%DATE% %TIME%] PYTHON_REINTENTO intento=%%I espera=%ESPERA_REINTENTO%s
        set /a "PINGS_ESPERA=ESPERA_REINTENTO+1"
        ping 127.0.0.1 -n !PINGS_ESPERA! -w 1000 >nul
    )
)

>>"%RUTA_LOG%" echo [%DATE% %TIME%] ERROR: Python no encontrado tras %MAX_INTENTOS% intentos
exit /b 1

:python_encontrado
>>"%RUTA_LOG%" echo [%DATE% %TIME%] Python encontrado ^| %PYTHON_BOOTSTRAP% %PYTHON_ARGUMENTOS%
if defined REGISTRO_LAPTOP_SOLO_BUSCAR exit /b 0
>>"%RUTA_LOG%" echo [%DATE% %TIME%] bootstrap solicitado
start "" /b "%PYTHON_BOOTSTRAP%" %PYTHON_ARGUMENTOS% "%~dp0bootstrap_laptop.pyw"
if errorlevel 1 (
    >>"%RUTA_LOG%" echo [%DATE% %TIME%] ERROR al solicitar bootstrap ^| codigo=!ERRORLEVEL!
    exit /b !ERRORLEVEL!
)
>>"%~dp0update.log" echo [%DATE% %TIME%] [ARRANQUE] bootstrap solicitado. Inicio del bat: %INICIO_BATCH%
exit /b 0

:buscar_python
set "INTENTO_ACTUAL=%~1"
call :probar_archivo venv_pythonw "%~dp0.venv\Scripts\pythonw.exe" ""
call :probar_archivo venv_python "%~dp0.venv\Scripts\python.exe" ""
call :probar_path PATH_pythonw pythonw.exe ""
call :probar_path PATH_python python.exe ""
call :probar_path PATH_pyw pyw.exe "-3"
call :probar_path PATH_py py.exe "-3"
call :probar_archivo launcher_local_pyw "%RAIZ_LOCALAPPDATA%\Programs\Python\Launcher\pyw.exe" "-3"
call :probar_archivo launcher_local_py "%RAIZ_LOCALAPPDATA%\Programs\Python\Launcher\py.exe" "-3"
call :probar_archivo launcher_windows_pyw "%RAIZ_WINDOWS%\pyw.exe" "-3"
call :probar_archivo launcher_windows_py "%RAIZ_WINDOWS%\py.exe" "-3"
call :probar_patron localappdata_pythonw "%RAIZ_LOCALAPPDATA%\Programs\Python\Python*" pythonw.exe ""
call :probar_patron localappdata_python "%RAIZ_LOCALAPPDATA%\Programs\Python\Python*" python.exe ""
call :probar_patron programfiles_pythonw "%RAIZ_PROGRAMFILES%\Python*" pythonw.exe ""
call :probar_patron programfiles_python "%RAIZ_PROGRAMFILES%\Python*" python.exe ""
call :probar_patron programfilesx86_pythonw "%RAIZ_PROGRAMFILES_X86%\Python*" pythonw.exe ""
call :probar_patron programfilesx86_python "%RAIZ_PROGRAMFILES_X86%\Python*" python.exe ""
exit /b 0

:probar_archivo
if defined PYTHON_BOOTSTRAP exit /b 0
if exist "%~2" (
    set "PYTHON_BOOTSTRAP=%~2"
    set "PYTHON_ARGUMENTOS=%~3"
    >>"%RUTA_LOG%" echo [%DATE% %TIME%] PYTHON_BUSQUEDA intento=%INTENTO_ACTUAL% %~1 SI ^| %~2
) else (
    >>"%RUTA_LOG%" echo [%DATE% %TIME%] PYTHON_BUSQUEDA intento=%INTENTO_ACTUAL% %~1 NO
)
exit /b 0

:probar_path
if defined PYTHON_BOOTSTRAP exit /b 0
set "RUTA_ENCONTRADA="
for %%D in ("!PATH:;=" "!") do if not defined RUTA_ENCONTRADA if exist "%%~D\%~2" set "RUTA_ENCONTRADA=%%~D\%~2"
if defined RUTA_ENCONTRADA (
    set "PYTHON_BOOTSTRAP=!RUTA_ENCONTRADA!"
    set "PYTHON_ARGUMENTOS=%~3"
    >>"%RUTA_LOG%" echo [%DATE% %TIME%] PYTHON_BUSQUEDA intento=%INTENTO_ACTUAL% %~1 SI ^| !RUTA_ENCONTRADA!
) else (
    >>"%RUTA_LOG%" echo [%DATE% %TIME%] PYTHON_BUSQUEDA intento=%INTENTO_ACTUAL% %~1 NO
)
exit /b 0

:probar_patron
if defined PYTHON_BOOTSTRAP exit /b 0
set "RUTA_ENCONTRADA="
for /d %%D in ("%~2") do if not defined RUTA_ENCONTRADA if exist "%%~fD\%~3" set "RUTA_ENCONTRADA=%%~fD\%~3"
if defined RUTA_ENCONTRADA (
    set "PYTHON_BOOTSTRAP=!RUTA_ENCONTRADA!"
    set "PYTHON_ARGUMENTOS=%~4"
    >>"%RUTA_LOG%" echo [%DATE% %TIME%] PYTHON_BUSQUEDA intento=%INTENTO_ACTUAL% %~1 SI ^| !RUTA_ENCONTRADA!
) else (
    >>"%RUTA_LOG%" echo [%DATE% %TIME%] PYTHON_BUSQUEDA intento=%INTENTO_ACTUAL% %~1 NO
)
exit /b 0
