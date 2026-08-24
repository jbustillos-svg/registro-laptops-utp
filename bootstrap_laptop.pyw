"""Prepara el entorno local y arranca la aplicación usando solo la biblioteca estándar."""

from datetime import datetime
import ast
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tkinter as tk


DIRECTORIO_APP = Path(__file__).resolve().parent
RUTA_VENV = DIRECTORIO_APP / ".venv"
RUTA_PYTHON = RUTA_VENV / "Scripts" / "python.exe"
RUTA_PYTHONW = RUTA_VENV / "Scripts" / "pythonw.exe"
RUTA_REQUIREMENTS = DIRECTORIO_APP / "requirements..txt"
RUTA_MARCADOR = DIRECTORIO_APP / ".requirements.sha256"
RUTA_LOG = DIRECTORIO_APP / "bootstrap.log"
RUTA_APLICACION = DIRECTORIO_APP / "registro_laptop.pyw"
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def registrar(mensaje):
    try:
        with RUTA_LOG.open("a", encoding="utf-8") as archivo:
            marca = datetime.now().isoformat(timespec="seconds")
            archivo.write(f"[{marca}] {mensaje}\n")
    except OSError:
        pass


class AvisoPreparacion:
    def __init__(self):
        self.raiz = None
        self.texto = None
        try:
            self.raiz = tk.Tk()
            self.raiz.title("Sistema de Control de Laptops")
            self.raiz.resizable(False, False)
            self.raiz.attributes("-topmost", True)
            self.texto = tk.Label(
                self.raiz,
                text=(
                    "Preparando el Sistema de Control de Laptops...\n\n"
                    "Este proceso se realiza únicamente cuando es necesario."
                ),
                font=("Segoe UI", 11),
                padx=32,
                pady=28,
                justify=tk.CENTER,
            )
            self.texto.pack()
            self.raiz.update_idletasks()
            ancho = self.raiz.winfo_reqwidth()
            alto = self.raiz.winfo_reqheight()
            x = (self.raiz.winfo_screenwidth() - ancho) // 2
            y = (self.raiz.winfo_screenheight() - alto) // 2
            self.raiz.geometry(f"{ancho}x{alto}+{x}+{y}")
            self.raiz.update()
        except tk.TclError as error:
            registrar(f"aviso no disponible: {error}")
            self.raiz = None

    def error(self):
        if not self.raiz or not self.texto:
            return
        self.texto.config(
            text=(
                "No fue posible preparar el sistema.\n\n"
                "Verifica la conexión a Internet e intenta nuevamente."
            )
        )
        boton = tk.Button(self.raiz, text="ENTENDIDO", command=self.raiz.destroy)
        boton.pack(pady=(0, 20), ipadx=18, ipady=5)
        self.raiz.update_idletasks()
        self.raiz.mainloop()
        self.raiz = None

    def cerrar(self):
        if self.raiz:
            try:
                self.raiz.destroy()
            except tk.TclError:
                pass
            self.raiz = None


def ejecutar(comando, timeout):
    return subprocess.run(
        comando,
        cwd=str(DIRECTORIO_APP),
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        timeout=timeout,
        creationflags=CREATE_NO_WINDOW,
        check=False,
    )


def hash_requirements():
    return hashlib.sha256(RUTA_REQUIREMENTS.read_bytes()).hexdigest()


def verificar_imports_entorno():
    modulos = set()
    for ruta in (RUTA_APLICACION, DIRECTORIO_APP / "proximidad_carrito.py"):
        try:
            arbol = ast.parse(ruta.read_text(encoding="utf-8-sig"))
        except (OSError, SyntaxError):
            return False
        for nodo in ast.walk(arbol):
            if isinstance(nodo, ast.Import):
                modulos.update(alias.name.split(".", 1)[0] for alias in nodo.names)
            elif isinstance(nodo, ast.ImportFrom) and nodo.module:
                modulos.add(nodo.module.split(".", 1)[0])

    codigo = (
        "import importlib\n"
        f"mods={json.dumps(sorted(modulos))}\n"
        "[importlib.import_module(nombre) for nombre in mods]\n"
    )
    resultado = ejecutar([str(RUTA_PYTHON), "-c", codigo], timeout=60)
    return resultado.returncode == 0


def entorno_preparado(hash_actual):
    if not RUTA_PYTHON.exists() or not RUTA_PYTHONW.exists():
        return False
    try:
        marcador_valido = (
            RUTA_MARCADOR.read_text(encoding="ascii").strip() == hash_actual
        )
    except OSError:
        return False
    return marcador_valido and verificar_imports_entorno()


def intentar_actualizacion_recuperacion():
    """Actualiza solo un árbol limpio cuando el entorno necesita preparación."""
    entorno = os.environ.copy()
    entorno["GIT_TERMINAL_PROMPT"] = "0"
    entorno["GCM_INTERACTIVE"] = "Never"

    def git(argumentos, timeout):
        return subprocess.run(
            ["git", *argumentos],
            cwd=str(DIRECTORIO_APP),
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            env=entorno,
            creationflags=CREATE_NO_WINDOW,
            check=False,
        )

    try:
        estado = git(["status", "--porcelain"], 5)
        if estado.returncode != 0 or estado.stdout.strip():
            return
        fetch = git(["fetch", "origin", "main", "--quiet"], 20)
        if fetch.returncode != 0:
            registrar("recuperación Git no disponible")
            return
        avance = git(["merge-base", "--is-ancestor", "HEAD", "origin/main"], 5)
        if avance.returncode != 0:
            registrar("recuperación Git omitida: no es avance rápido")
            return
        merge = git(["merge", "--ff-only", "origin/main"], 15)
        registrar(
            "recuperación Git aplicada"
            if merge.returncode == 0
            else "recuperación Git no aplicada"
        )
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as error:
        registrar(f"recuperación Git omitida: {error}")


def preparar_entorno(hash_actual):
    venv_existente = RUTA_PYTHON.exists() and RUTA_PYTHONW.exists()
    if not venv_existente:
        registrar("creando entorno virtual")
        resultado = ejecutar(
            [sys.executable, "-m", "venv", str(RUTA_VENV)],
            timeout=120,
        )
        if resultado.returncode != 0 or not RUTA_PYTHON.exists():
            registrar("error creando entorno virtual")
            return False, False

    registrar("instalando requirements")
    resultado = ejecutar(
        [str(RUTA_PYTHON), "-m", "pip", "install", "-r", str(RUTA_REQUIREMENTS)],
        timeout=600,
    )
    if resultado.returncode != 0:
        registrar("error instalando requirements")
        return False, venv_existente and verificar_imports_entorno()

    verificacion = ejecutar([str(RUTA_PYTHON), "-m", "pip", "check"], timeout=60)
    if verificacion.returncode != 0:
        registrar("error verificando requirements")
        return False, venv_existente and verificar_imports_entorno()

    try:
        RUTA_MARCADOR.write_text(hash_actual + "\n", encoding="ascii")
    except OSError as error:
        registrar(f"error guardando marcador: {error}")
        return False, venv_existente

    registrar("requirements preparados")
    return True, True


def iniciar_aplicacion():
    try:
        subprocess.Popen(
            [str(RUTA_PYTHONW), str(RUTA_APLICACION)],
            cwd=str(DIRECTORIO_APP),
            close_fds=True,
            creationflags=CREATE_NO_WINDOW,
        )
        registrar("aplicación iniciada")
        return True
    except OSError as error:
        registrar(f"error iniciando aplicación: {error}")
        return False


def main():
    aviso = None
    try:
        if not RUTA_REQUIREMENTS.exists() or not RUTA_APLICACION.exists():
            registrar("faltan archivos requeridos")
            aviso = AvisoPreparacion()
            aviso.error()
            return

        hash_actual = hash_requirements()
        if not entorno_preparado(hash_actual):
            aviso = AvisoPreparacion()
            intentar_actualizacion_recuperacion()
            hash_actual = hash_requirements()
            preparado, entorno_utilizable = preparar_entorno(hash_actual)
            if not preparado and not entorno_utilizable:
                aviso.error()
                return
            if not preparado:
                registrar("se usará el entorno anterior tras fallo de actualización")

        if aviso:
            aviso.cerrar()
        if not iniciar_aplicacion():
            aviso = AvisoPreparacion()
            aviso.error()
    except Exception as error:
        registrar(f"error inesperado: {type(error).__name__}: {error}")
        if aviso:
            aviso.error()
        else:
            AvisoPreparacion().error()


if __name__ == "__main__":
    main()
