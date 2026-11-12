"""Cliente de la cola local hacia el coordinador central Apps Script.

No escribe Google Sheets directamente. El Apps Script es el único proceso que
aplica operaciones críticas y resuelve la idempotencia global por UUID.
"""
import json
import os
import msvcrt
import re
import threading
from contextlib import contextmanager
from datetime import datetime
from urllib.parse import urlsplit

import requests


TIMEOUT_CONEXION = 3
TIMEOUT_RESPUESTA = 12
TIMEOUT_INTERACTIVO = (2, 5)
TIMEOUT_SEGUNDO_PLANO = (TIMEOUT_CONEXION, TIMEOUT_RESPUESTA)
TIMEOUT_PRE_APAGADO = (2, 3)
BACKOFF_DEPENDENCIA_INICIAL_SEGUNDOS = 5 * 60
BACKOFF_DEPENDENCIA_MAXIMO_SEGUNDOS = 6 * 60 * 60
TIPOS_CON_DEPENDENCIA_PRESTAMO = frozenset(
    ("DEVOLUCION", "NO_ENTREGA", "CIERRE_AUTOMATICO")
)


class ErrorCoordinador(RuntimeError):
    """Error público, con contexto técnico seguro sólo para el log local."""

    def __init__(self, codigo, diagnostico=None, api_key=None):
        super().__init__(codigo)
        self.diagnostico = diagnostico or {}
        self._api_key = api_key


def url_segura(url):
    """Devuelve únicamente esquema, host y ruta; nunca query, fragmento o clave."""
    try:
        partes = urlsplit(str(url))
        if not partes.scheme or not partes.hostname:
            return "URL_INVALIDA"
        host = partes.hostname
        try:
            if partes.port:
                host = f"{host}:{partes.port}"
        except ValueError:
            return "URL_INVALIDA"
        return f"{partes.scheme}://{host}{partes.path or '/'}"
    except (TypeError, ValueError):
        return "URL_INVALIDA"


def _texto_sanitizado(valor, api_key=None, limite=300):
    """Compacta texto de red y elimina secretos frecuentes antes de registrarlo."""
    texto = str(valor or "").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    if api_key:
        texto = texto.replace(str(api_key), "[REDACTADO]")
    texto = re.sub(
        r"(?i)([\"']?(?:api[_-]?key|authorization|access[_-]?token|id[_-]?token|token|password|cookie)[\"']?\s*[:=]\s*[\"']?)([^\"'\s,&}]+)",
        r"\1[REDACTADO]", texto,
    )
    texto = re.sub(r"(?i)\b(bearer\s+)[^\s,;]+", r"\1[REDACTADO]", texto)
    texto = re.sub(r"(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", "[EMAIL_REDACTADO]", texto)
    return texto[:limite]


def _cadena_redirecciones(respuesta):
    historial = list(getattr(respuesta, "history", None) or [])
    if not historial:
        return "ninguna"
    pasos = []
    for previo in historial:
        solicitud = getattr(previo, "request", None)
        metodo = getattr(solicitud, "method", "?")
        destino = ""
        cabeceras = getattr(previo, "headers", {}) or {}
        if hasattr(cabeceras, "get"):
            destino = cabeceras.get("Location", "")
        pasos.append(
            f"{getattr(previo, 'status_code', '?')}:{metodo}:{url_segura(getattr(previo, 'url', ''))}"
            + (f"->{url_segura(destino)}" if destino else "")
        )
    solicitud_final = getattr(respuesta, "request", None)
    metodo_final = getattr(solicitud_final, "method", "?")
    pasos.append(f"final:{metodo_final}:{url_segura(getattr(respuesta, 'url', ''))}")
    return ">".join(pasos)


def diagnostico_respuesta_coordinador(respuesta, api_key=None, incluir_muestra=False):
    """Metadatos seguros de la respuesta, aptos sólo para inicio_tecnico.log."""
    cabeceras = getattr(respuesta, "headers", {}) or {}
    tipo_contenido = cabeceras.get("Content-Type", "") if hasattr(cabeceras, "get") else ""
    detalle = {
        "http": getattr(respuesta, "status_code", "?"),
        "url": url_segura(getattr(respuesta, "url", "")),
        "content_type": _texto_sanitizado(tipo_contenido, api_key, 120) or "SIN_CONTENT_TYPE",
        "redirecciones": _cadena_redirecciones(respuesta),
    }
    if incluir_muestra:
        try:
            muestra = getattr(respuesta, "text", "")
        except Exception as error:  # Un mock o stream no debe ocultar el diagnóstico.
            muestra = f"LECTURA_CUERPO_{type(error).__name__}"
        detalle["muestra"] = _texto_sanitizado(muestra, api_key)
    return detalle


def diagnostico_error_coordinador(error):
    """Resume causa y contexto sin revelar API_KEY ni contenido funcional."""
    api_key = getattr(error, "_api_key", None)
    partes = [
        f"exc={type(error).__name__}",
        f"mensaje={_texto_sanitizado(error, api_key, 180)}",
    ]
    causa = error.__cause__
    if causa is not None:
        partes.append(
            f"causa={type(causa).__name__}:{_texto_sanitizado(causa, api_key, 180)}"
        )
    for clave, valor in getattr(error, "diagnostico", {}).items():
        partes.append(f"{clave}={_texto_sanitizado(valor, api_key, 350)}")
    return " ".join(partes)


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
    payload = {
        "api_key": api_key, "id_evento": evento["id_evento"],
        "tipo_evento": evento["tipo_evento"],
        "matricula": evento.get("matricula") or "", "nombre": evento.get("nombre") or "",
        "laptop_id": evento.get("laptop_id") or "", "equipo_origen": evento.get("equipo") or "",
        "fecha_evento": evento.get("fecha_evento") or "", "hora_evento": evento.get("hora_evento") or "",
        "fecha_creacion": evento.get("fecha_creacion") or "", "datos": datos,
    }
    try:
        # Google Apps Script suele devolver una redirección tras ejecutar doPost.
        # Se declara explícitamente para poder diagnosticar su cadena en la respuesta.
        respuesta = (post or requests.post)(
            url, json=payload, timeout=timeout, allow_redirects=True,
            headers={"Accept": "application/json", "Content-Type": "application/json"},
        )
    except requests.RequestException as error:
        raise ErrorCoordinador(
            "SOLICITUD_COORDINADOR_FALLIDA",
            {"url": url_segura(url), "causa": type(error).__name__},
            api_key,
        ) from error
    if getattr(respuesta, "status_code", 200) != 200:
        raise ErrorCoordinador(
            f"HTTP_{getattr(respuesta, 'status_code', 'ERROR')}",
            diagnostico_respuesta_coordinador(respuesta, api_key, incluir_muestra=True),
            api_key,
        )
    try:
        contenido = respuesta.json()
    except (ValueError, TypeError) as error:
        raise ErrorCoordinador(
            "RESPUESTA_COORDINADOR_INVALIDA",
            diagnostico_respuesta_coordinador(respuesta, api_key, incluir_muestra=True),
            api_key,
        ) from error
    if isinstance(contenido, dict) and contenido.get("estado") in ("APLICADO", "YA_APLICADO"):
        return contenido.get("referencia_registro") or "COORDINADOR"
    codigo = contenido.get("codigo") if isinstance(contenido, dict) else None
    raise ErrorCoordinador(
        str(codigo or "ERROR_COORDINADOR")[:120],
        diagnostico_respuesta_coordinador(respuesta, api_key),
        api_key,
    )


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

    def _enviar(self, evento, timeout=TIMEOUT_SEGUNDO_PLANO):
        return enviar_evento_coordinador(
            evento, self.configuracion, self._post, timeout
        )

    @staticmethod
    def _datos_evento(evento):
        try:
            datos = json.loads(evento.get("datos_json") or "{}")
        except (TypeError, ValueError):
            return {}
        return datos if isinstance(datos, dict) else {}

    @classmethod
    def _prestamo_dependencia_id(cls, evento):
        """Obtiene el UUID del pr\u00e9stamo de la misma sesi\u00f3n, si existe."""
        return str(cls._datos_evento(evento).get("prestamo_evento_id") or "").strip()

    @staticmethod
    def _orden_evento(evento):
        """Da salida a pr\u00e9stamos antes que a sus operaciones dependientes."""
        tipo = evento.get("tipo_evento")
        intentos = int(evento.get("intentos") or 0)
        if tipo == "PRESTAMO" and intentos == 0:
            prioridad = 0
        elif tipo == "PRESTAMO":
            prioridad = 1
        elif intentos == 0:
            prioridad = 2
        else:
            prioridad = 3
        return (prioridad, evento.get("fecha_creacion") or "", evento.get("id_evento") or "")

    @staticmethod
    def _instante(valor):
        try:
            return datetime.fromisoformat(str(valor))
        except (TypeError, ValueError):
            return None

    @classmethod
    def _en_backoff(cls, evento, ahora):
        """Difiere s\u00f3lo la dependencia remota repetidamente no encontrada."""
        if "DEPENDENCIA_PRESTAMO_NO_ENCONTRADA" not in str(evento.get("ultimo_error") or ""):
            return False
        ultimo = cls._instante(evento.get("fecha_ultimo_intento"))
        if ultimo is None:
            return False
        intentos = max(1, int(evento.get("intentos") or 0))
        espera = min(
            BACKOFF_DEPENDENCIA_INICIAL_SEGUNDOS * (2 ** (intentos - 1)),
            BACKOFF_DEPENDENCIA_MAXIMO_SEGUNDOS,
        )
        try:
            return (ahora - ultimo).total_seconds() < espera
        except TypeError:
            return False

    def sincronizar_evento(self, id_evento, timeout=TIMEOUT_PRE_APAGADO):
        """Intenta una sola vez un evento durable concreto, sin cambiar su UUID u hora."""
        if not self._lock.acquire(blocking=False):
            return False
        try:
            evento = self.almacen.obtener_evento(id_evento)
            if not evento:
                return False
            if evento.get("sincronizado"):
                return True
            if not self.configuracion.get("URL_APPS_SCRIPT") or not self.configuracion.get("API_KEY"):
                return False
            with self._bloqueo_interproceso() as adquirido:
                if not adquirido:
                    return False
                try:
                    referencia = self._enviar(evento, timeout)
                    self.almacen.marcar_sincronizado(evento["id_evento"], referencia)
                    return True
                except Exception as error:
                    self.almacen.incrementar_intentos(evento["id_evento"], str(error)[:500])
                    self.registrar(
                        "SINCRONIZACION_EVENTO_ERROR "
                        f"id={evento['id_evento']} tipo={evento['tipo_evento']} "
                        f"{diagnostico_error_coordinador(error)}"
                    )
                    return False
        finally:
            self._lock.release()

    def sincronizar(self, limite=30):
        """Confirma SQLite sólo después de una respuesta final del coordinador."""
        if not self._lock.acquire(blocking=False):
            return 0
        realizados = 0
        try:
            pendientes = self.almacen.obtener_eventos_pendientes()
            self.registrar(f"SINCRONIZACION_COLA eventos={len(pendientes)}")
            if not pendientes:
                return 0
            if not self.configuracion.get("URL_APPS_SCRIPT") or not self.configuracion.get("API_KEY"):
                return 0
            with self._bloqueo_interproceso() as adquirido:
                if not adquirido:
                    return 0
                ahora = datetime.now().astimezone()
                por_id = {evento["id_evento"]: evento for evento in pendientes}
                procesados = 0
                for evento in sorted(pendientes, key=self._orden_evento):
                    if procesados >= limite:
                        break
                    id_evento = evento["id_evento"]
                    if self._en_backoff(evento, ahora):
                        self.registrar(
                            "SINCRONIZACION_EVENTO_DIFERIDO "
                            f"id={id_evento} motivo=BACKOFF"
                        )
                        continue
                    prestamo_id = self._prestamo_dependencia_id(evento)
                    prestamo_pendiente = por_id.get(prestamo_id)
                    if (
                        evento.get("tipo_evento") in TIPOS_CON_DEPENDENCIA_PRESTAMO
                        and prestamo_pendiente
                        and not prestamo_pendiente.get("sincronizado")
                    ):
                        self.registrar(
                            "SINCRONIZACION_EVENTO_DIFERIDO "
                            f"id={id_evento} motivo=DEPENDENCIA_PRESTAMO_PENDIENTE"
                        )
                        continue
                    try:
                        referencia = self._enviar(evento)
                        self.almacen.marcar_sincronizado(id_evento, referencia)
                        evento["sincronizado"] = 1
                        realizados += 1
                        procesados += 1
                    except Exception as error:
                        self.almacen.incrementar_intentos(id_evento, str(error)[:500])
                        procesados += 1
                        self.registrar(
                            "SINCRONIZACION_EVENTO_ERROR "
                            f"id={id_evento} tipo={evento['tipo_evento']} "
                            f"{diagnostico_error_coordinador(error)}"
                        )
                        # No monopolizar el bloqueo con una cola histórica en
                        # error: el siguiente ciclo recuperará ese UUID.
                        self.registrar(
                            "SINCRONIZACION_EVENTO_CONTINUA_TRAS_ERROR "
                            f"id={id_evento}"
                        )
        finally:
            self._lock.release()
        pendientes_restantes = len(self.almacen.obtener_eventos_pendientes())
        self.registrar(
            "SINCRONIZACION_COMPLETA "
            f"aplicados={realizados} pendientes={pendientes_restantes}"
        )
        return realizados
