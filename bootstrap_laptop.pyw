"""Prepara el entorno local y arranca la aplicación usando solo la biblioteca estándar."""

from datetime import datetime
import ast
import hashlib
import json
import os
from pathlib import Path
import queue
import secrets
import socket
import subprocess
import sys
import tkinter as tk
from tkinter import ttk
import threading
import time


DIRECTORIO_APP = Path(__file__).resolve().parent
RUTA_VENV = DIRECTORIO_APP / ".venv"
RUTA_PYTHON = RUTA_VENV / "Scripts" / "python.exe"
RUTA_PYTHONW = RUTA_VENV / "Scripts" / "pythonw.exe"
RUTA_REQUIREMENTS = DIRECTORIO_APP / "requirements..txt"
RUTA_MARCADOR = DIRECTORIO_APP / ".requirements.sha256"
RUTA_LOG = DIRECTORIO_APP / "bootstrap.log"
RUTA_LOG_ARRANQUE = DIRECTORIO_APP / "arranque.log"
RUTA_APLICACION = DIRECTORIO_APP / "registro_laptop.pyw"
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
TIEMPO_ESPERA_APP_LISTA = 120


def registrar(mensaje):
    marca = datetime.now().isoformat(timespec="seconds")
    for ruta_log in (RUTA_LOG, RUTA_LOG_ARRANQUE):
        try:
            with ruta_log.open("a", encoding="utf-8") as archivo:
                archivo.write(f"[{marca}] {mensaje}\n")
        except OSError:
            pass


class AvisoPreparacion:
    def __init__(self):
        self.raiz = None
        self.texto = None
        self.barra = None
        self.cola = queue.Queue()
        self.error_definitivo = False
        try:
            self.raiz = tk.Tk()
            self.raiz.title("Sistema de Control de Laptops")
            self.raiz.attributes("-topmost", True)
            self.raiz.state("zoomed")
            self.raiz.overrideredirect(True)
            self.raiz.configure(bg="#0b1f3a")
            self.raiz.protocol("WM_DELETE_WINDOW", self._cierre_bloqueado)
            self.raiz.bind_all("<Alt-F4>", self._cierre_bloqueado)
            self.raiz.bind("<Unmap>", self._restaurar_pantalla)
            contenedor = tk.Frame(self.raiz, bg="#0b1f3a")
            contenedor.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            tk.Label(
                contenedor,
                text="SISTEMA DE CONTROL DE LAPTOPS",
                font=("Segoe UI", 25, "bold"),
                fg="white",
                bg="#0b1f3a",
            ).pack(pady=(0, 32))
            self.texto = tk.Label(
                contenedor,
                text="Cargando sistema...",
                font=("Segoe UI", 14),
                fg="#dce8f7",
                bg="#0b1f3a",
                justify=tk.CENTER,
            )
            self.texto.pack(pady=(0, 24))
            self.barra = ttk.Progressbar(
                contenedor,
                mode="indeterminate",
                length=420,
            )
            self.barra.pack(pady=(0, 38))
            self.barra.start(12)
            acciones = tk.Frame(contenedor, bg="#0b1f3a")
            acciones.pack()
            estilo_boton = {
                "font": ("Segoe UI", 11, "bold"),
                "fg": "white",
                "bg": "#244b75",
                "activeforeground": "white",
                "activebackground": "#326496",
                "relief": tk.FLAT,
                "width": 14,
                "pady": 10,
            }
            tk.Button(
                acciones,
                text="APAGAR",
                command=lambda: self._accion_windows("apagar"),
                **estilo_boton,
            ).pack(side=tk.LEFT, padx=10)
            tk.Button(
                acciones,
                text="REINICIAR",
                command=lambda: self._accion_windows("reiniciar"),
                **estilo_boton,
            ).pack(side=tk.LEFT, padx=10)
            self.raiz.focus_force()
        except tk.TclError as error:
            registrar(f"aviso no disponible: {error}")
            self.raiz = None

    def error(self):
        self.cola.put(("error", None))

    def actualizar(self, mensaje):
        self.cola.put(("mensaje", mensaje))

    def cerrar(self):
        self.cola.put(("cerrar", None))

    def _cierre_bloqueado(self):
        if self.texto:
            self.texto.config(
                text="El sistema todavía está iniciando..."
            )
        return "break"

    def _restaurar_pantalla(self, _evento=None):
        if self.raiz:
            self.raiz.after_idle(self._mostrar_al_frente)

    def _mostrar_al_frente(self):
        if not self.raiz:
            return
        try:
            self.raiz.deiconify()
            self.raiz.attributes("-topmost", True)
            self.raiz.lift()
            self.raiz.focus_force()
        except tk.TclError:
            pass

    def _accion_windows(self, accion):
        argumentos = ["shutdown", "/s" if accion == "apagar" else "/r", "/t", "0"]
        registrar(f"acción solicitada: {accion}")
        try:
            subprocess.Popen(argumentos, creationflags=CREATE_NO_WINDOW)
        except OSError as error:
            registrar(f"error solicitando {accion}: {error}")
            if self.texto:
                self.texto.config(text="No fue posible completar la acción solicitada.")

    def _destruir(self):
        if not self.raiz:
            return
        try:
            self.raiz.destroy()
        except tk.TclError:
            pass
        self.raiz = None

    def _mostrar_error(self):
        if not self.raiz or not self.texto:
            return
        self.error_definitivo = True
        if self.barra:
            self.barra.stop()
            self.barra.pack_forget()
        self.texto.config(
            text=(
                "No fue posible iniciar el sistema.\n\n"
                "Puedes apagar o reiniciar el equipo."
            )
        )

    def _procesar_cola(self):
        if not self.raiz:
            return
        try:
            while True:
                accion, valor = self.cola.get_nowait()
                if accion == "mensaje" and self.texto:
                    self.texto.config(text=valor)
                elif accion == "error":
                    self._mostrar_error()
                elif accion == "cerrar":
                    self._destruir()
                    return
        except queue.Empty:
            pass
        if self.raiz:
            self.raiz.after(100, self._procesar_cola)

    def ejecutar(self):
        if not self.raiz:
            return
        self.raiz.after(50, self._procesar_cola)
        self.raiz.mainloop()


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


def red_disponible(timeout=3):
    try:
        with socket.create_connection(("github.com", 443), timeout=timeout):
            return True
    except OSError:
        return False


def esperar_red(aviso, intervalo=10):
    registrar("red no disponible; esperando")
    aviso.actualizar("Esperando conexión...")
    while not red_disponible():
        for _ in range(intervalo):
            time.sleep(1)
            aviso.actualizar("Esperando conexión...")
    registrar("red disponible")
    aviso.actualizar("Preparando el Sistema de Control de Laptops...")


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
        estado = git(
            ["status", "--porcelain", "--untracked-files=no"], 5
        )
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


def crear_senal_app_lista():
    token = secrets.token_hex(16)
    directorio_temporal = Path(os.environ.get("TEMP", str(DIRECTORIO_APP)))
    ruta = directorio_temporal / f"registro_laptop_lista_{os.getpid()}_{token}.tmp"
    try:
        ruta.unlink(missing_ok=True)
    except OSError as error:
        registrar(f"no se pudo limpiar señal APP_LISTA nueva: {error}")
    return ruta, token


def iniciar_aplicacion(ruta_senal, token_senal):
    try:
        registrar("registro_laptop solicitado")
        entorno = os.environ.copy()
        entorno["REGISTRO_LAPTOP_APP_LISTA"] = str(ruta_senal)
        entorno["REGISTRO_LAPTOP_APP_TOKEN"] = token_senal
        proceso = subprocess.Popen(
            [str(RUTA_PYTHONW), str(RUTA_APLICACION)],
            cwd=str(DIRECTORIO_APP),
            close_fds=True,
            creationflags=CREATE_NO_WINDOW,
            env=entorno,
        )
        registrar("aplicación iniciada")
        return proceso
    except OSError as error:
        registrar(f"error iniciando aplicación: {error}")
        return None


def esperar_app_lista(proceso, ruta_senal, token_senal, timeout=TIEMPO_ESPERA_APP_LISTA):
    registrar(f"esperando APP_LISTA | timeout={timeout}s")
    limite = time.monotonic() + timeout
    while time.monotonic() < limite:
        try:
            if ruta_senal.read_text(encoding="ascii").strip() == token_senal:
                registrar("APP_LISTA recibida")
                return True
        except FileNotFoundError:
            pass
        except OSError as error:
            registrar(f"error leyendo APP_LISTA: {error}")
        if proceso.poll() is not None:
            registrar(
                f"registro_laptop finalizó antes de APP_LISTA | codigo={proceso.returncode}"
            )
            return False
        time.sleep(0.1)
    registrar("timeout esperando APP_LISTA")
    return False


def ejecutar_preparacion(aviso):
    """Realiza la preparación sin bloquear el hilo principal de Tkinter."""
    try:
        registrar("worker iniciado")
        if not RUTA_REQUIREMENTS.exists() or not RUTA_APLICACION.exists():
            registrar("faltan archivos requeridos")
            aviso.error()
            return

        hash_actual = hash_requirements()
        if entorno_preparado(hash_actual):
            registrar("venv válida")
            registrar(
                "red disponible" if red_disponible() else "red no disponible"
            )
        else:
            registrar("venv no válida")
            aviso.actualizar("Preparando el Sistema de Control de Laptops...")
            entorno_anterior_funcional = (
                RUTA_PYTHON.exists()
                and RUTA_PYTHONW.exists()
                and verificar_imports_entorno()
            )
            hay_red = red_disponible()
            if not hay_red:
                registrar("red no disponible")
                if entorno_anterior_funcional:
                    registrar("venv anterior funcional; se inicia sin actualizar")
                else:
                    esperar_red(aviso)
            else:
                registrar("red disponible")
            if not entorno_anterior_funcional or hay_red:
                intentar_actualizacion_recuperacion()
                hash_actual = hash_requirements()
                preparado = False
                entorno_utilizable = False
                for intento in range(1, 4):
                    registrar(f"preparación de entorno intento={intento}")
                    preparado, entorno_utilizable = preparar_entorno(hash_actual)
                    if preparado or entorno_utilizable:
                        break
                    if not red_disponible():
                        esperar_red(aviso)
                    elif intento < 3:
                        aviso.actualizar(
                            "Preparando el Sistema de Control de Laptops..."
                        )
                        time.sleep(10)
                if not preparado and not entorno_utilizable:
                    aviso.error()
                    return
                if not preparado:
                    registrar(
                        "se usará el entorno anterior tras fallo de actualización"
                    )

        registrar("preparación completada")
        aviso.actualizar("Iniciando sistema...")
        time.sleep(0.5)
        ruta_senal, token_senal = crear_senal_app_lista()
        proceso = iniciar_aplicacion(ruta_senal, token_senal)
        if not proceso:
            aviso.error()
            return
        if not esperar_app_lista(proceso, ruta_senal, token_senal):
            aviso.error()
            return
        registrar("bootstrap finalizado")
        aviso.cerrar()
    except Exception as error:
        registrar(f"error inesperado: {type(error).__name__}: {error}")
        aviso.error()
    finally:
        if "ruta_senal" in locals():
            try:
                ruta_senal.unlink(missing_ok=True)
            except OSError as error:
                registrar(f"error eliminando APP_LISTA: {error}")


def main():
    registrar(f"bootstrap iniciado | ejecutable={sys.executable}")
    aviso = AvisoPreparacion()
    if not aviso.raiz:
        ejecutar_preparacion(aviso)
        return

    worker = threading.Thread(
        target=ejecutar_preparacion,
        args=(aviso,),
        name="PreparacionBootstrap",
        daemon=False,
    )
    worker.start()
    aviso.ejecutar()


if __name__ == "__main__":
    main()
