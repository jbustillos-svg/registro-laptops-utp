"""Cliente de la cola local hacia el coordinador central Apps Script.

No escribe Google Sheets directamente. El Apps Script es el único proceso que
aplica operaciones críticas y resuelve la idempotencia global por UUID.
"""
import json
import os
import msvcrt
import threading
from contextlib import contextmanager

import requests


TIMEOUT_CONEXION = 3
TIMEOUT_RESPUESTA = 12
TIMEOUT_INTERACTIVO = (2, 5)
TIMEOUT_SEGUNDO_PLANO = (TIMEOUT_CONEXION, TIMEOUT_RESPUESTA)


class ErrorCoordinador(RuntimeError):
    pass


def enviar_evento_coordinador(evento, configuracion, post=None,
                              timeout=TIMEOUT_INTERACTIVO):
    """Única capa HTTP para eventos online y para reintentos de la cola."""
    url = (configuracion or {}).get("URL_APPS_SCRIPT")
    api_key = (configuracion or {}).get("API_KEY")
    if not url or not api_key:
        raise ErrorCoordinador("CONFIGURACION_SINCRONIZACION_NO_DISPONIBLE")
    try:
        datos = json.loads(evento.get("datos_json") or "{}")
    except (TypeError, ValueError) as error:
        raise ErrorCoordinador("DATOS_EVENTO_INVALIDOS") from error
    respuesta = (post or requests.post)(url, json={
        "api_key": api_key, "id_evento": evento["id_evento"],
        "tipo_evento": evento["tipo_evento"],
        "matricula": evento.get("matricula") or "", "nombre": evento.get("nombre") or "",
        "laptop_id": evento.get("laptop_id") or "", "equipo_origen": evento.get("equipo") or "",
        "fecha_evento": evento.get("fecha_evento") or "", "hora_evento": evento.get("hora_evento") or "",
        "fecha_creacion": evento.get("fecha_creacion") or "", "datos": datos,
    }, timeout=timeout)
    if getattr(respuesta, "status_code", 200) != 200:
        raise ErrorCoordinador(f"HTTP_{getattr(respuesta, 'status_code', 'ERROR')}")
    try:
        contenido = respuesta.json()
    except (ValueError, TypeError) as error:
        raise ErrorCoordinador("RESPUESTA_COORDINADOR_INVALIDA") from error
    if isinstance(contenido, dict) and contenido.get("estado") in ("APLICADO", "YA_APLICADO"):
        return contenido.get("referencia_registro") or "COORDINADOR"
    codigo = contenido.get("codigo") if isinstance(contenido, dict) else None
    raise ErrorCoordinador(str(codigo or "ERROR_COORDINADOR")[:120])


class SincronizadorContingencia:
    def __init__(self, almacen, configuracion, registrar, post=None):
        self.almacen = almacen
        self.configuracion = configuracion or {}
        self.registrar = registrar
        self._post = post or requests.post
        self._lock = threading.Lock()

    @contextmanager
    def _bloqueo_interproceso(self):
        """Evita dos workers de la misma laptop; no sustituye el lock central."""
        ruta = str(self.almacen.ruta_db) + ".sync.lock"
        with open(ruta, "a+b") as archivo:
            archivo.seek(0, os.SEEK_END)
            if archivo.tell() == 0:
                archivo.write(b"0")
                archivo.flush()
            archivo.seek(0)
            try:
                msvcrt.locking(archivo.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError:
                yield False
                return
            try:
                yield True
            finally:
                archivo.seek(0)
                msvcrt.locking(archivo.fileno(), msvcrt.LK_UNLCK, 1)

    @staticmethod
    def _payload(evento, api_key):
        try:
            datos = json.loads(evento.get("datos_json") or "{}")
        except (TypeError, ValueError) as error:
            raise ValueError("DATOS_EVENTO_INVALIDOS") from error
        return {
            "api_key": api_key,
            "id_evento": evento["id_evento"],
            "tipo_evento": evento["tipo_evento"],
            "matricula": evento.get("matricula") or "",
            "nombre": evento.get("nombre") or "",
            "laptop_id": evento.get("laptop_id") or "",
            "equipo_origen": evento.get("equipo") or "",
            "fecha_evento": evento.get("fecha_evento") or "",
            "hora_evento": evento.get("hora_evento") or "",
            "fecha_creacion": evento.get("fecha_creacion") or "",
            "datos": datos,
        }

    def _enviar(self, evento):
        return enviar_evento_coordinador(
            evento, self.configuracion, self._post, TIMEOUT_SEGUNDO_PLANO
        )

    def sincronizar(self, limite=30):
        """Confirma SQLite sólo después de una respuesta final del coordinador."""
        if not self._lock.acquire(blocking=False):
            return 0
        realizados = 0
        try:
            pendientes = self.almacen.obtener_eventos_pendientes()
            if not pendientes:
                return 0
            if not self.configuracion.get("URL_APPS_SCRIPT") or not self.configuracion.get("API_KEY"):
                return 0
            with self._bloqueo_interproceso() as adquirido:
                if not adquirido:
                    return 0
                for evento in pendientes[:limite]:
                    try:
                        referencia = self._enviar(evento)
                        self.almacen.marcar_sincronizado(evento["id_evento"], referencia)
                        realizados += 1
                    except Exception as error:
                        self.almacen.incrementar_intentos(evento["id_evento"], str(error)[:500])
                        self.registrar(
                            "SINCRONIZACION_EVENTO_ERROR "
                            f"id={evento['id_evento']} tipo={evento['tipo_evento']} "
                            f"error={type(error).__name__}"
                        )
        finally:
            self._lock.release()
        return realizados
