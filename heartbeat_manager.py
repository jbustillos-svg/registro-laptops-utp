"""Heartbeat opcional y no bloqueante hacia Laptop Manager."""

import json
import os
import socket
import threading
import urllib.error
import urllib.request


ARCHIVO_CONFIGURACION = "manager_config.json"
INTERVALO_SEGUNDOS = 30
TIMEOUT_SEGUNDOS = 3

_worker_lock = threading.Lock()
_worker = None


def cargar_configuracion(directorio_app):
    """Carga una configuración habilitada o devuelve None silenciosamente."""
    ruta = os.path.join(directorio_app, ARCHIVO_CONFIGURACION)
    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            config = json.load(archivo)
    except (OSError, ValueError, TypeError):
        return None

    if not isinstance(config, dict) or config.get("habilitado") is not True:
        return None

    url = config.get("url")
    api_key = config.get("api_key")
    if not isinstance(url, str) or not url.strip():
        return None
    if not isinstance(api_key, str) or not api_key:
        return None

    return {"url": url.rstrip("/"), "api_key": api_key}


class HeartbeatManager:
    def __init__(
        self,
        config,
        version,
        registrar_evento,
        intervalo=INTERVALO_SEGUNDOS,
        timeout=TIMEOUT_SEGUNDOS,
        hostname=None,
        urlopen=None,
    ):
        self._config = config
        self._version = version
        self._registrar_evento = registrar_evento
        self._intervalo = intervalo
        self._timeout = timeout
        self._hostname = hostname or socket.gethostname()
        self._urlopen = urlopen or urllib.request.urlopen
        self._detener = threading.Event()
        self._thread = threading.Thread(
            target=self._ejecutar,
            name="HeartbeatManager",
            daemon=True,
        )

    @property
    def thread(self):
        return self._thread

    def iniciar(self):
        self._thread.start()

    def _registrar(self, evento):
        try:
            self._registrar_evento(evento)
        except Exception:
            pass

    def _enviar(self):
        contenido = json.dumps(
            {"equipo": self._hostname, "version": self._version}
        ).encode("utf-8")
        solicitud = urllib.request.Request(
            f"{self._config['url']}/api/heartbeat",
            data=contenido,
            headers={
                "Content-Type": "application/json",
                "X-Laptop-Manager-Key": self._config["api_key"],
            },
            method="POST",
        )
        try:
            with self._urlopen(solicitud, timeout=self._timeout) as respuesta:
                estado = respuesta.getcode()
            if estado == 200:
                self._registrar("HEARTBEAT_MANAGER enviado")
            else:
                self._registrar(f"HEARTBEAT_MANAGER no disponible HTTP_{estado}")
        except urllib.error.HTTPError as error:
            self._registrar(
                f"HEARTBEAT_MANAGER no disponible HTTP_{error.code}"
            )
        except Exception as error:
            self._registrar(
                f"HEARTBEAT_MANAGER no disponible {type(error).__name__}"
            )

    def _ejecutar(self):
        self._registrar("HEARTBEAT_MANAGER iniciado")
        while not self._detener.is_set():
            self._enviar()
            self._detener.wait(self._intervalo)


def iniciar_heartbeat_manager(
    directorio_app,
    version,
    registrar_evento,
    intervalo=INTERVALO_SEGUNDOS,
    timeout=TIMEOUT_SEGUNDOS,
):
    """Inicia como máximo un worker; devuelve None si está deshabilitado."""
    config = cargar_configuracion(directorio_app)
    if config is None:
        return None

    global _worker
    with _worker_lock:
        if _worker is not None and _worker.thread.is_alive():
            return _worker
        _worker = HeartbeatManager(
            config,
            version,
            registrar_evento,
            intervalo=intervalo,
            timeout=timeout,
        )
        _worker.iniciar()
        return _worker
