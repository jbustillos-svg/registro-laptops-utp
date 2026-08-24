import ctypes
from datetime import datetime
import os
from pathlib import Path
import subprocess
import sys


CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def registrar_log(directorio, mensaje):
    ruta_log = Path(directorio) / "update.log"
    marca_tiempo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with ruta_log.open("a", encoding="utf-8") as archivo:
            archivo.write(f"[{marca_tiempo}] [ACTUALIZACION] {mensaje}\n")
    except OSError:
        pass


def esperar_fin_proceso(pid):
    synchronize = 0x00100000
    infinite = 0xFFFFFFFF
    handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, pid)
    if handle:
        try:
            ctypes.windll.kernel32.WaitForSingleObject(handle, infinite)
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)


def ejecutar_git(directorio, argumentos, timeout=15):
    entorno = os.environ.copy()
    entorno["GIT_TERMINAL_PROMPT"] = "0"
    entorno["GCM_INTERACTIVE"] = "Never"
    return subprocess.run(
        ["git", *argumentos],
        cwd=directorio,
        stdin=subprocess.DEVNULL,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=entorno,
        creationflags=CREATE_NO_WINDOW,
        check=False
    )


def main():
    if len(sys.argv) != 3:
        return

    pid_aplicacion = int(sys.argv[1])
    directorio = os.path.abspath(sys.argv[2])
    esperar_fin_proceso(pid_aplicacion)

    try:
        estado = ejecutar_git(
            directorio,
            ["status", "--porcelain", "--untracked-files=no"]
        )
        if estado.returncode != 0:
            registrar_log(directorio, "error comprobando cambios locales al aplicar")
            return
        if estado.stdout.strip():
            registrar_log(
                directorio,
                "pendiente: existen cambios locales versionados; no se aplicó"
            )
            return

        resultado = ejecutar_git(
            directorio,
            ["merge", "--ff-only", "origin/main"]
        )
        if resultado.returncode == 0:
            registrar_log(directorio, "actualización aplicada para el siguiente inicio")
            return

        detalle = (resultado.stderr or resultado.stdout or "error desconocido")
        detalle = " ".join(detalle.split())[:300]
        registrar_log(directorio, f"error al aplicar: {detalle}")
    except FileNotFoundError:
        registrar_log(directorio, "error: Git no está disponible")
    except subprocess.TimeoutExpired:
        registrar_log(directorio, "error: tiempo agotado al aplicar")
    except Exception as error:
        registrar_log(directorio, f"error inesperado al aplicar: {error}")


if __name__ == "__main__":
    main()
