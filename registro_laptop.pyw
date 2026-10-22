import tkinter as tk
import threading
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import socket
import requests
import pytz
import os
import time
import traceback
import hashlib
import hmac
import subprocess
import sys
import certifi
import tempfile
import msvcrt
import uuid
from contextlib import contextmanager
from contingencia_local import (
    CACHE_NO_DISPONIBLE,
    CACHE_OK,
    INTERNET_ERROR,
    INTERNET_OK,
    SHEETS_API_ERROR,
    SHEETS_AUTH_ERROR,
    SHEETS_OK,
    SHEETS_TIMEOUT,
    ContingenciaLocal,
)
from sincronizador_contingencia import (
    SincronizadorContingencia,
)
from sincronizacion_config import (
    actualizar_configuracion_desde_hoja,
    cargar_configuracion_local,
)
from proximidad_carrito import (
    ESTADO_ERROR,
    ESTADO_NO_DETECTADO,
    verificar_proximidad_carrito
)
from heartbeat_manager import iniciar_heartbeat_manager
from identidad_roles import (
    normalizar_rol,
    rol_requiere_curp,
    seleccionar_flujo_identidad,
)

DIRECTORIO_APP = os.path.dirname(os.path.abspath(__file__))
RUTA_CREDENCIALES = os.path.join(DIRECTORIO_APP, "credenciales.json")
RUTA_UTP = os.path.join(DIRECTORIO_APP, "UTP.png")

# Intentar importar PIL, si no está disponible usar emojis
try:
    from PIL import Image, ImageTk, ImageOps, ImageChops
    PIL_DISPONIBLE = True
except ImportError:
    PIL_DISPONIBLE = False

# --- VARIABLES GLOBALES ---
VERSION_SISTEMA = "1.4.9"
MODO_PRUEBA = False
PROXIMIDAD_HABILITADA = True
hoja_alumnos = None
hoja_registros = None
hoja_configuracion_sistema = None
zona_horaria = pytz.timezone("America/Chihuahua")
verificar_conexion_activo = True
procesando_sesion = False
aviso_internet = None
acepta_estado_equipo = None
chk_label = None
verificacion_aviso_en_curso = False
conexion_google_lock = threading.Lock()
actualizacion_lock = threading.Lock()
sincronizacion_lock = threading.Lock()
configuracion_sincronizacion_lock = threading.Lock()
dialogo_admin_abierto = False
dialogo_admin_referencia = None
dialogo_admin_estado = None
contexto_contingencia_proximidad = None
confirmacion_contingencia_abierta = False
atajo_administrativo_en_curso = False
ultimo_atajo_administrativo = 0.0
registro_contingencias_sesion = []
ventana_entrega_activa = None
almacen_contingencia = None
diagnostico_conectividad = {
    "internet": None,
    "sheets": None,
    "cache": CACHE_NO_DISPONIBLE,
}

TIEMPO_INACTIVIDAD_MS = 60000
TIMEOUT_SINCRONIZACION_PRE_APAGADO = (2, 3)
TIMEOUT_SINCRONIZACION_INMEDIATA = (2, 3)
ESPERA_LOCK_SINCRONIZACION_INMEDIATA = 0.25
timeout_inactividad_id = None
timeout_inactividad_activo = False
flujo_cancelado_por_inactividad = False
ventanas_flujo_alumno = set()

HASH_CLAVE_ADMIN = "0da417ade5c1bd3c262b600330f15e11af9b87a5b3984caac3d136a307b63190"
ARCHIVO_LOG_CONTINGENCIA = os.path.join(
    DIRECTORIO_APP, "contingencia_admin.log"
)
ARCHIVO_LOG_ACTUALIZACION = "update.log"
ARCHIVO_LOG_INICIO = "inicio_tecnico.log"
cierre_aplicacion_en_curso = False


class ErrorConsultaGoogle(Exception):
    """Indica que Google Sheets no pudo entregar un resultado confiable."""


def registrar_evento_tecnico(evento):
    directorio = os.path.dirname(os.path.abspath(__file__))
    ruta_log = os.path.join(directorio, ARCHIVO_LOG_INICIO)
    try:
        with open(ruta_log, "a", encoding="utf-8") as archivo:
            archivo.write(
                f"[{datetime.now().astimezone().isoformat(timespec='seconds')}] "
                f"evento={evento}\tpid={os.getpid()}\thostname={socket.gethostname()}\n"
            )
    except OSError:
        pass


def actualizar_diagnostico_conectividad(componente, estado):
    """Registra estados independientes para red, Sheets y caché local."""
    if diagnostico_conectividad.get(componente) == estado:
        return
    diagnostico_conectividad[componente] = estado
    registrar_evento_tecnico(f"CONECTIVIDAD componente={componente} estado={estado}")


def clasificar_error_sheets(error):
    if isinstance(error, (requests.Timeout, socket.timeout)):
        return SHEETS_TIMEOUT
    detalle = f"{type(error).__name__}: {error}".lower()
    if any(texto in detalle for texto in ("auth", "credential", "token", "permission", "forbidden", "unauthorized")):
        return SHEETS_AUTH_ERROR
    if "timeout" in detalle or "timed out" in detalle:
        return SHEETS_TIMEOUT
    return SHEETS_API_ERROR


def inicializar_infraestructura_contingencia():
    """Prepara SQLite al arrancar, sin impedir el funcionamiento actual si falla."""
    global almacen_contingencia
    try:
        almacen_contingencia = ContingenciaLocal(DIRECTORIO_APP)
        ruta = almacen_contingencia.inicializar()
        estado = almacen_contingencia.estado_cache()
        actualizar_diagnostico_conectividad("cache", estado)
        registrar_evento_tecnico(
            f"CONTINGENCIA_DB_LISTA ruta={ruta} schema_version="
            f"{almacen_contingencia.version_esquema()}"
        )
    except Exception as error:
        almacen_contingencia = None
        actualizar_diagnostico_conectividad("cache", CACHE_NO_DISPONIBLE)
        registrar_evento_tecnico(
            f"CONTINGENCIA_DB_ERROR={type(error).__name__}: {error}"
        )


def actualizar_cache_persona(matricula, **datos):
    """Guarda sólo datos confirmados por consultas normales de Sheets."""
    if almacen_contingencia is None:
        return
    try:
        almacen_contingencia.upsert_persona(matricula, **datos)
        actualizar_diagnostico_conectividad("cache", CACHE_OK)
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_CACHE_ERROR={type(error).__name__}: {error}"
        )


def obtener_persona_cache(matricula):
    if almacen_contingencia is None:
        return None
    try:
        return almacen_contingencia.obtener_persona(matricula)
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_CACHE_LECTURA_ERROR={type(error).__name__}: {error}"
        )
        return None


def _fecha_hora_local():
    ahora = datetime.now(zona_horaria)
    return ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S")


def enviar_evento_local_inmediato(id_evento):
    """Despierta la cola persistente sin esperar un POST desde el flujo de UI."""
    if almacen_contingencia is None:
        return False
    evento = almacen_contingencia.obtener_evento(id_evento)
    if not evento:
        return False
    solicitar_sincronizacion_inmediata(evento["tipo_evento"], id_evento=id_evento)
    # La solicitud puede omitirse porque un worker ya procesa la misma cola.
    # El evento sigue durable y pendiente hasta que el sincronizador lo confirme.
    return True


def sincronizar_evento_inmediato(id_evento, timeout=TIMEOUT_SINCRONIZACION_INMEDIATA):
    """Intenta una vez el UUID durable mediante el sincronizador existente.

    Debe invocarse desde un worker. La cola y el ciclo de 45 segundos siguen
    siendo el respaldo si no hay red, el coordinador falla o ya hay un ciclo
    procesando la misma base local.
    """
    if almacen_contingencia is None:
        return False
    try:
        evento = almacen_contingencia.obtener_evento(id_evento)
    except Exception as error:
        registrar_evento_tecnico(
            f"SINCRONIZACION_INMEDIATA_ERROR tipo={type(error).__name__}"
        )
        return False
    if not evento or evento.get("sincronizado"):
        return bool(evento and evento.get("sincronizado"))

    registrar_evento_tecnico(
        f"SINCRONIZACION_INMEDIATA_SOLICITADA tipo={evento['tipo_evento']}"
    )
    if not sincronizacion_lock.acquire(
        timeout=ESPERA_LOCK_SINCRONIZACION_INMEDIATA
    ):
        registrar_evento_tecnico(
            "SINCRONIZACION_INMEDIATA_OMITIDA motivo=ya_en_ejecucion"
        )
        return False
    try:
        registrar_evento_tecnico("SINCRONIZACION_INMEDIATA_INICIADA")
        configuracion = cargar_configuracion_local(DIRECTORIO_APP)
        aplicada = SincronizadorContingencia(
            almacen_contingencia, configuracion, registrar_evento_tecnico
        ).sincronizar_evento(id_evento, timeout=timeout)
        registrar_evento_tecnico(
            "SINCRONIZACION_INMEDIATA_APLICADA"
            if aplicada else "SINCRONIZACION_INMEDIATA_PENDIENTE"
        )
        return aplicada
    except Exception as error:
        registrar_evento_tecnico(
            f"SINCRONIZACION_INMEDIATA_ERROR tipo={type(error).__name__}"
        )
        return False
    finally:
        sincronizacion_lock.release()


def crear_y_enviar_evento_critico(tipo, matricula, nombre, fecha, hora, laptop_id,
                                  datos, clave_estable):
    """Persistencia primero: nunca hay fallback de escritura crítica por gspread."""
    if almacen_contingencia is None:
        return False
    id_evento = str(uuid.uuid5(uuid.NAMESPACE_URL, f"utp-v141:{tipo}:{clave_estable}"))
    almacen_contingencia.crear_evento(
        socket.gethostname(), tipo, matricula, nombre, fecha, hora, laptop_id,
        datos, id_evento=id_evento,
    )
    return enviar_evento_local_inmediato(id_evento)


def registrar_prestamo_contingencia(matricula, nombre):
    if almacen_contingencia is None:
        return False
    try:
        fecha, hora = _fecha_hora_local()
        sesion = almacen_contingencia.registrar_prestamo_offline(
            matricula, nombre, socket.gethostname(), fecha, hora
        )
        registrar_evento_tecnico(
            f"CONTINGENCIA_PRESTAMO_GUARDADO matricula={matricula} fecha={fecha} hora={hora}"
        )
        id_evento = sesion.get("prestamo_evento_id")
        if id_evento:
            enviar_evento_local_inmediato(id_evento)
        return True
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_PRESTAMO_ERROR={type(error).__name__}: {error}"
        )
        return False


def persistir_prestamo_online_local(matricula, nombre, fecha, hora, laptop_id):
    """No crea evento: Sheets ya confirmó el préstamo normal."""
    if almacen_contingencia is None:
        return
    try:
        almacen_contingencia.registrar_prestamo_online(
            matricula, nombre, laptop_id, fecha, hora
        )
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_SESION_ONLINE_ERROR={type(error).__name__}: {error}"
        )


def obtener_sesion_local_activa(matricula=None):
    if almacen_contingencia is None:
        return None
    try:
        sesion = almacen_contingencia.obtener_sesion_activa(socket.gethostname())
        if sesion and matricula and sesion["matricula"] != str(matricula).strip():
            return None
        if sesion:
            registrar_evento_tecnico(
                f"SESION_ACTIVA_CONSULTADA id_sesion={sesion['id_sesion']}"
            )
        return sesion
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_SESION_LECTURA_ERROR={type(error).__name__}: {error}"
        )
        return None


def registrar_sesion_pendiente_al_iniciar():
    """Cierra una sesión heredada; nunca la recupera para una nueva ejecución."""
    sesion = obtener_sesion_local_activa()
    if not sesion:
        registrar_evento_tecnico("SESION_RESIDUAL_DETECTADA resultado=NO_ACTIVA")
        return None

    registrar_evento_tecnico(
        f"SESION_RESIDUAL_DETECTADA id_sesion={sesion['id_sesion']}"
    )
    try:
        fecha, hora = _fecha_hora_local()
        resultado = almacen_contingencia.cerrar_sesion_residual_al_arrancar(
            socket.gethostname(), fecha, hora, socket.gethostname()
        )
        cerrada = resultado.get("sesion")
        if resultado.get("resultado") != "SESION_RESIDUAL_CERRADA" or not cerrada:
            registrar_evento_tecnico("SESION_RESIDUAL_CIERRE_OMITIDO motivo=SIN_SESION")
            return None
        if resultado.get("no_entrega_creada"):
            registrar_evento_tecnico(
                f"NO_ENTREGA_RESIDUAL_CREADA id_sesion={cerrada['id_sesion']}"
            )
        for campo in ("no_entrega_evento_id", "cierre_automatico_evento_id"):
            if cerrada.get(campo):
                enviar_evento_local_inmediato(cerrada[campo])
        registrar_evento_tecnico(
            f"SESION_RESIDUAL_CERRADA id_sesion={cerrada['id_sesion']}"
        )
        registrar_evento_tecnico("FLUJO_SELECCIONADO tipo=PRESTAMO_NUEVO")
        return cerrada
    except Exception as error:
        registrar_evento_tecnico(
            f"SESION_RESIDUAL_CIERRE_ERROR tipo={type(error).__name__}"
        )
        return None


def finalizar_devolucion_durable(matricula, solicitar_sincronizacion=True):
    """Única autoridad local: crea DEVOLUCION y cierra su sesión exacta."""
    if almacen_contingencia is None:
        return False
    try:
        sesion_previa = obtener_sesion_local_activa(matricula)
        if not sesion_previa:
            registrar_evento_tecnico(
                "SESION_LOCAL_CIERRE_IGNORADO id_sesion=NO_ENCONTRADA "
                "motivo=DEVOLUCION_SIN_SESION_ACTIVA"
            )
            return False
        registrar_evento_tecnico(
            f"DEVOLUCION_INICIADA id_sesion={sesion_previa['id_sesion']}"
        )
        fecha, hora = _fecha_hora_local()
        registrar_evento_tecnico(
            "SESION_LOCAL_CIERRE_SOLICITADO "
            f"id_sesion={sesion_previa['id_sesion']} tipo_evento=DEVOLUCION"
        )
        sesion = almacen_contingencia.registrar_devolucion_offline(
            matricula, socket.gethostname(), fecha, hora,
            id_sesion=sesion_previa["id_sesion"],
        )
        registrar_evento_tecnico(
            f"SESION_LOCAL_CERRADA id_sesion={sesion['id_sesion']}"
        )
        registrar_evento_tecnico(
            f"DEVOLUCION_EVENTO_DURABLE id={sesion.get('devolucion_evento_id')}"
        )
        id_evento = sesion.get("devolucion_evento_id")
        if id_evento and solicitar_sincronizacion:
            enviar_evento_local_inmediato(id_evento)
        return sesion
    except Exception as error:
        registrar_evento_tecnico(
            f"DEVOLUCION_DURABLE_ERROR={type(error).__name__}: {error}"
        )
        return False


def registrar_devolucion_contingencia(matricula, solicitar_sincronizacion=True):
    """Compatibilidad para llamadores existentes de la devolución durable."""
    return finalizar_devolucion_durable(matricula, solicitar_sincronizacion)


def persistir_devolucion_online_local(matricula, hora_salida):
    if almacen_contingencia is None:
        return
    try:
        sesion_previa = obtener_sesion_local_activa(matricula)
        if not sesion_previa:
            registrar_evento_tecnico(
                "SESION_LOCAL_CIERRE_IGNORADO id_sesion=NO_ENCONTRADA "
                "motivo=DEVOLUCION_ONLINE_SIN_SESION_ACTIVA"
            )
            return
        registrar_evento_tecnico(
            "SESION_LOCAL_CIERRE_SOLICITADO "
            f"id_sesion={sesion_previa['id_sesion']} tipo_evento=DEVOLUCION"
        )
        almacen_contingencia.registrar_devolucion_online(
            matricula, socket.gethostname(), hora_salida,
            id_sesion=sesion_previa["id_sesion"],
        )
        registrar_evento_tecnico(
            f"SESION_LOCAL_CERRADA id_sesion={sesion_previa['id_sesion']}"
        )
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_SALIDA_ONLINE_ERROR={type(error).__name__}: {error}"
        )


def no_entregas_efectivas_local(matricula, no_entregas_cache):
    if almacen_contingencia is None:
        return int(no_entregas_cache or 0)
    try:
        return almacen_contingencia.no_entregas_efectivas(matricula, no_entregas_cache)
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_NO_ENTREGAS_ERROR={type(error).__name__}: {error}"
        )
        return int(no_entregas_cache or 0)


def actualizar_configuracion_sincronizacion():
    """Lee la hoja central sólo al conectar o en revisiones poco frecuentes."""
    if hoja_configuracion_sistema is None:
        return cargar_configuracion_local(DIRECTORIO_APP)
    if not configuracion_sincronizacion_lock.acquire(blocking=False):
        return cargar_configuracion_local(DIRECTORIO_APP)
    try:
        configuracion, actualizada = actualizar_configuracion_desde_hoja(
            DIRECTORIO_APP, hoja_configuracion_sistema
        )
        if actualizada:
            registrar_evento_tecnico("SINCRONIZACION_CONFIG_ACTUALIZADA")
        return configuracion
    except Exception as error:
        registrar_evento_tecnico(
            f"SINCRONIZACION_CONFIG_LECTURA_ERROR={type(error).__name__}"
        )
        return cargar_configuracion_local(DIRECTORIO_APP)
    finally:
        configuracion_sincronizacion_lock.release()


def _sincronizar_pendientes_con_lock_adquirido():
    """Procesa la cola; el llamador conserva sincronizacion_lock durante todo el ciclo."""
    if not almacen_contingencia:
        return
    configuracion = cargar_configuracion_local(DIRECTORIO_APP)
    if not configuracion:
        return
    cantidad = SincronizadorContingencia(
        almacen_contingencia, configuracion, registrar_evento_tecnico
    ).sincronizar()
    if cantidad:
        registrar_evento_tecnico(f"SINCRONIZACION_COMPLETA eventos={cantidad}")


def sincronizar_pendientes_en_segundo_plano():
    if not sincronizacion_lock.acquire(blocking=False):
        return
    try:
        _sincronizar_pendientes_con_lock_adquirido()
    finally:
        sincronizacion_lock.release()


def solicitar_sincronizacion_inmediata(tipo_evento, id_evento=None):
    """Da prioridad al UUID nuevo sin abrir un segundo coordinador ni POST paralelo."""
    registrar_evento_tecnico(
        f"SINCRONIZACION_INMEDIATA_SOLICITADA tipo={tipo_evento}"
    )
    if not sincronizacion_lock.acquire(blocking=False):
        registrar_evento_tecnico(
            "SINCRONIZACION_INMEDIATA_OMITIDA motivo=ya_en_ejecucion"
        )
        return False

    def sincronizar():
        try:
            registrar_evento_tecnico("SINCRONIZACION_INMEDIATA_INICIADA")
            configuracion = cargar_configuracion_local(DIRECTORIO_APP)
            sincronizador = SincronizadorContingencia(
                almacen_contingencia, configuracion, registrar_evento_tecnico
            )
            if id_evento:
                sincronizador.sincronizar_evento(id_evento)
            else:
                sincronizador.sincronizar()
        finally:
            sincronizacion_lock.release()

    threading.Thread(target=sincronizar, daemon=True).start()
    return True


def sincronizar_devolucion_pre_apagado(id_evento):
    """Da una sola oportunidad breve a DEVOLUCION antes del shutdown.

    SQLite ya contiene el evento y su hora original al entrar aquí. Un fallo de
    red sólo conserva la cola pendiente: no cambia el resultado de la entrega.
    """
    if almacen_contingencia is None:
        return False
    registrar_evento_tecnico("DEVOLUCION_PRE_APAGADO_SINCRONIZACION_INICIADA")
    try:
        aplicada = sincronizar_evento_inmediato(
            id_evento, timeout=TIMEOUT_SINCRONIZACION_PRE_APAGADO
        )
        registrar_evento_tecnico(
            "DEVOLUCION_PRE_APAGADO_SINCRONIZACION_APLICADA"
            if aplicada else "DEVOLUCION_PRE_APAGADO_SINCRONIZACION_PENDIENTE"
        )
        return aplicada
    except Exception as error:
        registrar_evento_tecnico(
            "DEVOLUCION_PRE_APAGADO_SINCRONIZACION_ERROR="
            f"{type(error).__name__}"
        )
        return False


def iniciar_sincronizacion_pendientes():
    threading.Thread(target=sincronizar_pendientes_en_segundo_plano, daemon=True).start()


def cerrar_sesion_contingencia_por_nueva_sesion(matricula_nueva):
    """Resuelve en SQLite la sesión previa de otra matrícula y envía sus UUID."""
    if almacen_contingencia is None:
        return False
    try:
        fecha, hora = _fecha_hora_local()
        laptop_id = socket.gethostname()
        previa = obtener_sesion_local_activa()
        if not previa:
            registrar_evento_tecnico(
                "SESION_LOCAL_CIERRE_IGNORADO id_sesion=NO_ENCONTRADA "
                "motivo=CIERRE_AUTOMATICO_SIN_SESION_ACTIVA"
            )
            return True
        registrar_evento_tecnico(
            "SESION_LOCAL_CIERRE_SOLICITADO "
            f"id_sesion={previa['id_sesion']} tipo_evento=NO_ENTREGA"
        )
        registrar_evento_tecnico(
            "SESION_LOCAL_CIERRE_SOLICITADO "
            f"id_sesion={previa['id_sesion']} tipo_evento=CIERRE_AUTOMATICO"
        )
        resultado = almacen_contingencia.resolver_sesion_activa_por_nueva_matricula(
            laptop_id, matricula_nueva, fecha, hora, laptop_id,
            registrar_no_entrega=not matricula_exenta_no_entrega(previa["matricula"]),
        )
        sesion = resultado.get("sesion")
        if resultado["resultado"] != "SESION_ANTERIOR_CERRADA":
            registrar_evento_tecnico(
                "SESION_LOCAL_CIERRE_IGNORADO "
                f"id_sesion={previa['id_sesion']} motivo={resultado['resultado']}"
            )
            return resultado["resultado"] == "SIN_SESION"
        for campo in ("no_entrega_evento_id", "cierre_automatico_evento_id"):
            id_evento = sesion.get(campo)
            if id_evento:
                enviar_evento_local_inmediato(id_evento)
        registrar_evento_tecnico(
            "CONTINGENCIA_CIERRE_AUTOMATICO "
            f"matricula_anterior={sesion['matricula']} matricula_nueva={matricula_nueva}"
        )
        registrar_evento_tecnico(
            f"SESION_LOCAL_CERRADA id_sesion={sesion['id_sesion']}"
        )
        return True
    except Exception as error:
        registrar_evento_tecnico(
            f"CONTINGENCIA_CIERRE_AUTOMATICO_ERROR={type(error).__name__}: {error}"
        )
        return False


def limpiar_logs_antiguos(dias=7):
    """Elimina archivos .log antiguos ubicados directamente en DIRECTORIO_APP."""
    eliminados = 0
    limite_segundos = dias * 24 * 60 * 60
    ahora = time.time()

    try:
        with os.scandir(DIRECTORIO_APP) as entradas:
            for entrada in entradas:
                try:
                    if not entrada.is_file(follow_symlinks=False):
                        continue
                    if not entrada.name.lower().endswith(".log"):
                        continue
                    if ahora - entrada.stat(follow_symlinks=False).st_mtime <= limite_segundos:
                        continue
                    os.remove(entrada.path)
                    eliminados += 1
                except OSError:
                    continue
    except OSError:
        pass

    registrar_evento_tecnico(f"LIMPIEZA_LOGS eliminados={eliminados}")
    return eliminados


def iniciar_limpieza_logs():
    threading.Thread(target=limpiar_logs_antiguos, daemon=True).start()


def registrar_inicio_tecnico():
    directorio = os.path.dirname(os.path.abspath(__file__))
    commit = "NO_DISPONIBLE"
    try:
        resultado = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=directorio,
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=3,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            check=False
        )
        if resultado.returncode == 0:
            commit = resultado.stdout.strip() or commit
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass

    ruta_log = os.path.join(directorio, ARCHIVO_LOG_INICIO)
    try:
        with open(ruta_log, "a", encoding="utf-8") as archivo:
            archivo.write(
                f"[{datetime.now().astimezone().isoformat(timespec='seconds')}] "
                f"ruta={os.path.abspath(__file__)}\tpid={os.getpid()}\t"
                f"hostname={socket.gethostname()}\tcommit={commit}\n"
            )
    except OSError:
        pass

# =========================
# COLUMNAS HOJA REGISTROS (0-based)
# =========================
COL_MATRICULA = 0
COL_NOMBRE = 1
COL_FECHA = 2
COL_HORA_INGRESO = 3
COL_HORA_SALIDA = 4
COL_LAPTOP_ID = 5
COL_OBSERVACION = 6


def observacion_es_normal(valor):
    return (valor or "").strip() in ("", "S/N")


ESTADO_VACIO_VALIDO = "ESTADO_VACIO_VALIDO"


def normalizar_estado_alumno(valor, lectura_exitosa=False):
    if valor == ESTADO_VACIO_VALIDO:
        return ESTADO_VACIO_VALIDO
    if valor is None:
        return ESTADO_VACIO_VALIDO if lectura_exitosa else None

    estado = str(valor).strip().upper()
    if estado == "":
        return ESTADO_VACIO_VALIDO
    return estado if estado in ("ACTIVO", "ADVERTENCIA", "SANCIONADO") else None


def estado_alumno_permite_prestamo(valor):
    return normalizar_estado_alumno(valor) in {
        ESTADO_VACIO_VALIDO,
        "ACTIVO",
        "ADVERTENCIA",
    }


@contextmanager
def bloqueo_local_registro_prestamo():
    """Serializa el tramo leer/append entre procesos de esta laptop."""
    identificador = hashlib.sha256(
        os.path.abspath(__file__).encode("utf-8")
    ).hexdigest()[:16]
    ruta = os.path.join(tempfile.gettempdir(), f"registro_laptop_{identificador}.lock")
    with open(ruta, "a+b") as archivo:
        archivo.seek(0, os.SEEK_END)
        if archivo.tell() == 0:
            archivo.write(b"0")
            archivo.flush()
        archivo.seek(0)
        msvcrt.locking(archivo.fileno(), msvcrt.LK_LOCK, 1)
        try:
            yield
        finally:
            archivo.seek(0)
            msvcrt.locking(archivo.fileno(), msvcrt.LK_UNLCK, 1)

MATRICULAS_EXENTAS_NO_ENTREGA = {
    "010537700",
    "082025",
    "11",
}
# Cambiar a "clasico" permite recuperar la apariencia anterior sin tocar código.
TEMA_INTERFAZ = "moderno"

TEMAS_INTERFAZ = {
    "clasico": {
        "colores": {
            "primario": "#0066cc", "secundario": "#00a8ff",
            "fondo": "#f5f7fa", "tarjeta": "#ffffff",
            "texto": "#2d3436", "texto_secundario": "#636e72",
            "exito": "#00b894", "error": "#e84118",
            "advertencia": "#fdcb6e", "borde": "#dfe6e9",
            "hover": "#f1f2f6", "campo": "#f8f9fa",
            "info_suave": "#eef6ff", "advertencia_suave": "#fff4d6"
        },
        "fuentes": {
            "titulo": ("Segoe UI", 24, "bold"),
            "subtitulo": ("Segoe UI", 14, "bold"),
            "cuerpo": ("Segoe UI", 11),
            "cuerpo_bold": ("Segoe UI", 11, "bold"),
            "entrada": ("Segoe UI", 14),
            "boton": ("Segoe UI", 12, "bold"),
            "pequena": ("Segoe UI", 9),
            "pequena_bold": ("Segoe UI", 9, "bold")
        },
        "medidas": {"tarjeta_login_ancho": 460, "tarjeta_login_alto": 620,
                    "dialogo_ancho": 440, "radio_borde": 1},
        "marca_agua_fondo_login": False
    },
    "moderno": {
        "colores": {
            "primario": "#124E78", "secundario": "#0E7490",
            "fondo": "#EDF2F7", "tarjeta": "#FFFFFF",
            "texto": "#172B3A", "texto_secundario": "#5F7180",
            "exito": "#21835B", "error": "#C2413B",
            "advertencia": "#D89B20", "borde": "#D7E0E8",
            "hover": "#E7EEF4", "campo": "#F7F9FB",
            "info_suave": "#EAF3F8", "advertencia_suave": "#FFF6DE"
        },
        "fuentes": {
            "titulo": ("Segoe UI", 26, "bold"),
            "subtitulo": ("Segoe UI", 15, "bold"),
            "cuerpo": ("Segoe UI", 11),
            "cuerpo_bold": ("Segoe UI", 11, "bold"),
            "entrada": ("Segoe UI", 15),
            "boton": ("Segoe UI", 11, "bold"),
            "pequena": ("Segoe UI", 9),
            "pequena_bold": ("Segoe UI", 9, "bold")
        },
        "medidas": {"tarjeta_login_ancho": 480, "tarjeta_login_alto": 640,
                    "dialogo_ancho": 460, "radio_borde": 1},
        "marca_agua_fondo_login": True
    }
}

if TEMA_INTERFAZ not in TEMAS_INTERFAZ:
    TEMA_INTERFAZ = "clasico"

ESTILO_ACTUAL = TEMAS_INTERFAZ[TEMA_INTERFAZ]
COLORES = ESTILO_ACTUAL["colores"]
FUENTES = ESTILO_ACTUAL["fuentes"]
TAMANOS = ESTILO_ACTUAL["medidas"]
FONDO_MARCA_AGUA_LOGIN = ESTILO_ACTUAL["marca_agua_fondo_login"]

COLOR_PRIMARIO = COLORES["primario"]
COLOR_SECUNDARIO = COLORES["secundario"]
COLOR_FONDO = COLORES["fondo"]
COLOR_TARJETA = COLORES["tarjeta"]
COLOR_TEXTO = COLORES["texto"]
COLOR_TEXTO_SECUNDARIO = COLORES["texto_secundario"]
COLOR_EXITO = COLORES["exito"]
COLOR_ERROR = COLORES["error"]
COLOR_ADVERTENCIA = COLORES["advertencia"]
COLOR_BORDE = COLORES["borde"]
COLOR_HOVER = COLORES["hover"]

FUENTE_TITULO = FUENTES["titulo"]
FUENTE_SUBTITULO = FUENTES["subtitulo"]
FUENTE_CUERPO = FUENTES["cuerpo"]
FUENTE_CUERPO_BOLD = FUENTES["cuerpo_bold"]
FUENTE_ENTRADA = FUENTES["entrada"]
FUENTE_BOTON = FUENTES["boton"]
FUENTE_PEQ = FUENTES["pequena"]
FUENTE_PEQ_BOLD = FUENTES["pequena_bold"]


def estilo_boton(tipo="principal"):
    fondos = {
        "principal": COLOR_PRIMARIO,
        "secundario": COLOR_HOVER,
        "exito": COLOR_EXITO,
        "error": COLOR_ERROR
    }
    fondos_activos = {
        "principal": COLOR_SECUNDARIO,
        "secundario": COLOR_BORDE,
        "exito": "#176B49",
        "error": "#9F332F"
    }
    fondo = fondos.get(tipo, COLOR_PRIMARIO)
    return {
        "font": FUENTE_BOTON,
        "bg": fondo,
        "fg": COLOR_TEXTO if tipo == "secundario" else "white",
        "activebackground": fondos_activos.get(tipo, COLOR_SECUNDARIO),
        "activeforeground": COLOR_TEXTO if tipo == "secundario" else "white",
        "bd": 0,
        "relief": tk.FLAT,
        "cursor": "hand2"
    }


def estilo_entrada():
    return {
        "font": FUENTE_ENTRADA,
        "bg": COLORES["campo"],
        "fg": COLOR_TEXTO,
        "insertbackground": COLOR_TEXTO,
        "relief": tk.FLAT,
        "highlightthickness": 2,
        "highlightbackground": COLOR_BORDE,
        "highlightcolor": COLOR_PRIMARIO
    }


def _centrar_dialogo(
    dialogo,
    parent,
    ancho_minimo=440,
    alto_minimo=190,
    margen_horizontal=12,
    margen_vertical=12
):
    dialogo.update_idletasks()
    ancho_requerido = dialogo.winfo_reqwidth()
    alto_requerido = dialogo.winfo_reqheight()
    ancho = max(ancho_requerido + margen_horizontal, ancho_minimo)
    alto = max(alto_requerido + margen_vertical, alto_minimo)

    parent.update_idletasks()
    ancho_parent = parent.winfo_width()
    alto_parent = parent.winfo_height()

    if ancho_parent > 1 and alto_parent > 1:
        x = parent.winfo_rootx() + (ancho_parent - ancho) // 2
        y = parent.winfo_rooty() + (alto_parent - alto) // 2
    else:
        x = (dialogo.winfo_screenwidth() - ancho) // 2
        y = (dialogo.winfo_screenheight() - alto) // 2

    dialogo.geometry(f"{ancho}x{alto}+{x}+{y}")
    dialogo.deiconify()
    dialogo.lift()


def _crear_dialogo_personalizado(titulo, mensaje, tipo, parent, texto_destacado=None):
    estilos = {
        "info": (COLOR_PRIMARIO, "i"),
        "advertencia": ("#f0ad00", "!"),
        "error": (COLOR_ERROR, "×"),
        "exito": (COLOR_EXITO, "✓")
    }
    color, icono = estilos.get(tipo, estilos["info"])
    parent = parent or ventana

    dialogo = tk.Toplevel(parent)
    vincular_dispatcher_administrativo(dialogo)
    dialogo.title(titulo)
    dialogo.resizable(False, False)
    dialogo.configure(bg=COLOR_FONDO)
    dialogo.transient(parent)
    dialogo.attributes("-topmost", True)
    dialogo.withdraw()

    tarjeta = tk.Frame(
        dialogo,
        bg=COLOR_TARJETA,
        highlightthickness=1,
        highlightbackground=COLOR_BORDE
    )
    tarjeta.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    if TEMA_INTERFAZ == "moderno":
        tk.Frame(tarjeta, height=4, bg=color).pack(fill=tk.X, side=tk.TOP)

    botones = tk.Frame(tarjeta, bg=COLOR_TARJETA, padx=22)
    botones.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 18))

    cuerpo = tk.Frame(tarjeta, bg=COLOR_TARJETA, padx=22, pady=20)
    cuerpo.pack(fill=tk.BOTH, expand=True)

    icono_canvas = tk.Canvas(
        cuerpo,
        width=56,
        height=56,
        bg=COLOR_TARJETA,
        highlightthickness=0
    )
    icono_canvas.pack(side=tk.LEFT, anchor="n", padx=(0, 18))
    icono_canvas.create_oval(4, 4, 52, 52, fill=color, outline=color)
    icono_canvas.create_text(
        28,
        28,
        text=icono,
        fill="white",
        font=("Segoe UI", 21, "bold")
    )

    texto = tk.Frame(cuerpo, bg=COLOR_TARJETA)
    texto.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    tk.Label(
        texto,
        text=titulo,
        font=("Segoe UI", 15, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA,
        anchor="w"
    ).pack(fill=tk.X, pady=(2, 8))
    tk.Label(
        texto,
        text=mensaje,
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA,
        justify=tk.LEFT,
        anchor="w",
        wraplength=360
    ).pack(fill=tk.X)

    if texto_destacado:
        tk.Label(
            texto,
            text=texto_destacado,
            font=FUENTE_CUERPO_BOLD,
            fg="#8A5A00",
            bg=COLORES["advertencia_suave"],
            justify=tk.LEFT,
            anchor="w",
            wraplength=340,
            padx=12,
            pady=9
        ).pack(fill=tk.X, pady=(12, 0))

    return dialogo, botones, color, parent


def mostrar_dialogo_personalizado(
    titulo,
    mensaje,
    tipo="info",
    parent=None,
    al_crear=None,
    texto_destacado=None,
    texto_boton="ACEPTAR"
):
    timeout_iniciado_por_dialogo = False
    if procesando_sesion and not timeout_inactividad_activo:
        activar_timeout_inactividad()
        timeout_iniciado_por_dialogo = True

    dialogo, botones, color, parent = _crear_dialogo_personalizado(
        titulo, mensaje, tipo, parent, texto_destacado
    )
    registrar_ventana_flujo_alumno(dialogo)
    grab_anterior = parent.grab_current()

    def cerrar():
        try:
            dialogo.grab_release()
        except tk.TclError:
            pass
        dialogo.destroy()
        try:
            if grab_anterior and grab_anterior.winfo_exists():
                grab_anterior.grab_set()
        except tk.TclError:
            pass

    tk.Button(
        botones,
        text=texto_boton,
        font=FUENTE_BOTON,
        bg=color,
        fg="white",
        activebackground=color,
        activeforeground="white",
        bd=0,
        cursor="hand2",
        padx=26,
        pady=8,
        command=cerrar
    ).pack(side=tk.RIGHT)

    dialogo.protocol("WM_DELETE_WINDOW", cerrar)
    dialogo.bind("<Return>", lambda _evento: cerrar())
    dialogo.bind("<Escape>", lambda _evento: cerrar())
    if al_crear:
        al_crear(dialogo, cerrar)
    _centrar_dialogo(dialogo, parent, ancho_minimo=TAMANOS["dialogo_ancho"])
    dialogo.grab_set()
    dialogo.focus_force()
    dialogo.wait_window()
    if timeout_iniciado_por_dialogo:
        cancelar_timeout_inactividad()


def mostrar_confirmacion_personalizada(titulo, mensaje, parent=None):
    dialogo, botones, color, parent = _crear_dialogo_personalizado(
        titulo, mensaje, "info", parent
    )
    registrar_ventana_flujo_alumno(dialogo)
    grab_anterior = parent.grab_current()
    resultado = {"valor": False}

    def cerrar(valor=False):
        resultado["valor"] = valor
        try:
            dialogo.grab_release()
        except tk.TclError:
            pass
        dialogo.destroy()
        try:
            if grab_anterior and grab_anterior.winfo_exists():
                grab_anterior.grab_set()
        except tk.TclError:
            pass

    tk.Button(
        botones,
        text="NO",
        padx=26,
        pady=8,
        command=lambda: cerrar(False),
        **estilo_boton("secundario")
    ).pack(side=tk.RIGHT, padx=(12, 0))
    tk.Button(
        botones,
        text="SÍ",
        padx=26,
        pady=8,
        command=lambda: cerrar(True),
        **estilo_boton("principal")
    ).pack(side=tk.RIGHT)

    dialogo.protocol("WM_DELETE_WINDOW", lambda: cerrar(False))
    dialogo.bind("<Return>", lambda _evento: cerrar(True))
    dialogo.bind("<Escape>", lambda _evento: cerrar(False))
    _centrar_dialogo(dialogo, parent, ancho_minimo=TAMANOS["dialogo_ancho"])
    dialogo.grab_set()
    dialogo.focus_force()
    dialogo.wait_window()
    return resultado["valor"]


def mostrar_informacion(titulo, mensaje, parent=None):
    return mostrar_dialogo_personalizado(titulo, mensaje, "info", parent)


def mostrar_advertencia(titulo, mensaje, parent=None):
    return mostrar_dialogo_personalizado(titulo, mensaje, "advertencia", parent)


def mostrar_error(
    titulo,
    mensaje,
    parent=None,
    al_crear=None,
    texto_destacado=None,
    texto_boton="ACEPTAR"
):
    return mostrar_dialogo_personalizado(
        titulo,
        mensaje,
        "error",
        parent,
        al_crear,
        texto_destacado,
        texto_boton
    )


def mostrar_exito(titulo, mensaje, parent=None):
    return mostrar_dialogo_personalizado(titulo, mensaje, "exito", parent)


def mostrar_fallo_proximidad(estado, parent=None, al_crear=None):
    registrar_evento_tecnico(f"PROXIMIDAD_DIALOGO_ABIERTO estado={estado}")
    if estado == ESTADO_NO_DETECTADO:
        mostrar_error(
            "Carrito no detectado",
            "No fue posible detectar la señal del carrito móvil.\n"
            "Verifique que el carrito se encuentre encendido.",
            parent=parent,
            al_crear=al_crear,
            texto_destacado="SOLICITE APOYO AL RESPONSABLE"
        )
        return

    if estado == ESTADO_ERROR:
        mostrar_error(
            "Error de verificación",
            "No fue posible verificar la ubicación del equipo.\n"
            "Inténtelo nuevamente si el problema continúa.",
            parent=parent,
            al_crear=al_crear,
            texto_destacado="SOLICITE APOYO AL RESPONSABLE"
        )
        return

    mostrar_error(
        "Laptop fuera de la ubicación autorizada",
        "Para continuar, la laptop debe encontrarse dentro del área del carrito móvil.\n"
        "Acérquese al carrito e inténtelo nuevamente.",
        parent=parent,
        al_crear=al_crear,
        texto_destacado="SI NECESITA AYUDA, SOLICITE APOYO AL RESPONSABLE"
    )


def registrar_uso_contingencia(tipo_operacion, matricula):
    registro_contingencias_sesion.append({
        "fecha_hora": datetime.now(zona_horaria).isoformat(timespec="seconds"),
        "evento": "CONTINGENCIA_ADMIN",
        "operacion": tipo_operacion,
        "matricula": matricula or "NO_DISPONIBLE",
        "equipo": socket.gethostname()
    })
    registro = registro_contingencias_sesion[-1]
    evento = (
        f"{registro['fecha_hora']}\t{registro['evento']}\t{registro['operacion']}\t"
        f"matricula={registro['matricula']}\tequipo={registro['equipo']}\n"
    )

    if MODO_PRUEBA:
        print(evento.rstrip())
        return

    try:
        with open(ARCHIVO_LOG_CONTINGENCIA, "a", encoding="utf-8") as archivo:
            archivo.write(evento)
    except OSError as error:
        print(f"[CONTINGENCIA_ADMIN] No se pudo escribir el registro local: {error}")


def _widget_tk_vivo(widget):
    if widget is None:
        return False
    try:
        return bool(widget.winfo_exists())
    except (tk.TclError, AttributeError):
        return False


def _restaurar_grab_administrativo(grab_anterior, foco_anterior):
    """Devuelve control al modal previo si todavía existe."""
    if not _widget_tk_vivo(grab_anterior):
        return
    try:
        grab_anterior.grab_set()
        if _widget_tk_vivo(foco_anterior):
            foco_anterior.focus_set()
        else:
            grab_anterior.focus_force()
        registrar_evento_tecnico("ADMIN_GRAB_RESTAURADO")
    except tk.TclError:
        pass


def finalizar_dialogo_administrativo(motivo="CANCELADO", restaurar=True, dialogo=None):
    """Cierra idempotentemente el diálogo administrativo y libera su estado.

    También se invoca si Tk destruye el diálogo indirectamente al destruir el
    modal padre durante el timeout de inactividad.
    """
    global dialogo_admin_abierto, dialogo_admin_referencia, dialogo_admin_estado

    estado = dialogo_admin_estado
    if not estado or estado.get("finalizado"):
        return False
    if dialogo is not None and estado.get("dialogo") is not dialogo:
        return False

    estado["finalizado"] = True
    dialogo = estado.get("dialogo")
    for after_id in tuple(estado.get("after_pendientes", ())):
        try:
            dialogo.after_cancel(after_id)
        except (tk.TclError, AttributeError):
            pass
    estado["after_pendientes"] = []

    if _widget_tk_vivo(dialogo):
        try:
            dialogo.grab_release()
        except tk.TclError:
            pass
        try:
            dialogo.destroy()
        except tk.TclError:
            pass

    dialogo_admin_abierto = False
    dialogo_admin_referencia = None
    dialogo_admin_estado = None
    registrar_evento_tecnico(f"ADMIN_DIALOGO_CIERRE motivo={motivo}")
    registrar_evento_tecnico("ADMIN_ESTADO_LIMPIADO")
    if restaurar:
        _restaurar_grab_administrativo(
            estado.get("grab_anterior"), estado.get("foco_anterior")
        )
    registrar_evento_tecnico("ADMIN_REAPERTURA_PERMITIDA")
    return True


def abrir_acceso_administrativo(
    event=None,
    parent=None,
    contexto="DESCONOCIDO",
    accion="SALIR_SISTEMA",
    al_autorizar=None,
):
    global dialogo_admin_abierto, dialogo_admin_referencia, dialogo_admin_estado

    registrar_evento_tecnico(
        f"ADMIN_SOLICITUD estado_guard={dialogo_admin_abierto} "
        f"contexto={contexto} accion={accion}"
    )
    if dialogo_admin_abierto:
        return "break"

    if accion not in ("CONTINUAR_SIN_PROXIMIDAD", "SALIR_SISTEMA"):
        raise ValueError(f"Acción administrativa no válida: {accion}")

    parent = parent or ventana
    registrar_evento_tecnico(
        f"ADMIN_DIALOGO_SOLICITADO accion={accion} contexto={contexto}"
    )

    grab_anterior = None
    foco_anterior = None
    try:
        grab_anterior = ventana.grab_current()
        foco_anterior = ventana.focus_get()
    except tk.TclError:
        pass

    if _widget_tk_vivo(grab_anterior):
        try:
            grab_anterior.grab_release()
            registrar_evento_tecnico("ADMIN_GRAB_TEMPORAL_LIBERADO")
        except tk.TclError:
            grab_anterior = None

    try:
        dialogo = tk.Toplevel(parent)
    except Exception:
        dialogo_admin_abierto = False
        _restaurar_grab_administrativo(grab_anterior, foco_anterior)
        raise
    dialogo_admin_abierto = True
    dialogo_admin_referencia = dialogo
    dialogo_admin_estado = {
        "dialogo": dialogo,
        "grab_anterior": grab_anterior,
        "foco_anterior": foco_anterior,
        "after_pendientes": [],
        "finalizado": False,
    }
    es_bypass_proximidad = accion == "CONTINUAR_SIN_PROXIMIDAD"
    titulo = (
        "Autorización administrativa"
        if es_bypass_proximidad else "Cierre administrativo"
    )
    mensaje = (
        "Introduzca la contraseña administrativa para continuar sin "
        "verificación de proximidad."
        if es_bypass_proximidad else
        "Introduzca la contraseña administrativa para cerrar el sistema."
    )
    texto_boton = "CONTINUAR" if es_bypass_proximidad else "CERRAR SISTEMA"
    tipo_boton = "principal" if es_bypass_proximidad else "error"

    dialogo.title(titulo)
    dialogo.resizable(False, False)
    dialogo.configure(bg=COLOR_TARJETA)
    dialogo.transient(parent)
    dialogo.attributes("-topmost", True)
    dialogo.withdraw()

    contenido = tk.Frame(dialogo, bg=COLOR_TARJETA, padx=22, pady=18)
    contenido.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        contenido,
        text=titulo,
        font=FUENTE_SUBTITULO,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 10))

    mensaje_label = tk.Label(
        contenido,
        text=mensaje,
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA,
        wraplength=460
    )
    mensaje_label.pack(pady=(0, 14))

    tk.Label(
        contenido,
        text="Contraseña:",
        font=FUENTE_PEQ_BOLD,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(anchor="w")

    clave_var = tk.StringVar()
    entrada_clave = tk.Entry(
        contenido,
        textvariable=clave_var,
        show="*",
        justify=tk.CENTER,
        bd=0,
        width=24,
        **estilo_entrada()
    )
    entrada_clave.pack(anchor="w", ipady=8, pady=(5, 0))

    estado_intentos = tk.Label(
        contenido,
        text="",
        font=FUENTE_PEQ,
        fg=COLOR_ERROR,
        bg=COLOR_TARJETA
    )
    estado_intentos.pack(pady=(8, 0))

    botones = tk.Frame(contenido, bg=COLOR_TARJETA)
    botones.pack(fill=tk.X, pady=(14, 0))
    intentos_incorrectos = {"cantidad": 0}

    def cerrar(motivo="CANCELADO", restaurar=True):
        finalizar_dialogo_administrativo(motivo, restaurar, dialogo)

    def completar_autorizacion():
        global contexto_contingencia_proximidad
        global confirmacion_contingencia_abierta
        global ventana_entrega_activa

        registrar_evento_tecnico(f"ADMIN_AUTORIZACION_CORRECTA accion={accion}")
        cerrar("AUTORIZADO", restaurar=False)

        if accion == "CONTINUAR_SIN_PROXIMIDAD":
            if al_autorizar:
                al_autorizar()
            return

        registrar_evento_tecnico("SALIDA_ADMIN_AUTORIZADA")
        contexto_contingencia_proximidad = None
        confirmacion_contingencia_abierta = False
        ventana_entrega_activa = None
        cerrar_aplicacion()

    def autenticar():

        clave = clave_var.get()
        clave_var.set("")
        hash_ingresado = hashlib.sha256(clave.encode("utf-8")).hexdigest()
        clave = ""

        if hmac.compare_digest(hash_ingresado, HASH_CLAVE_ADMIN):
            completar_autorizacion()
            return

        intentos_incorrectos["cantidad"] += 1
        if intentos_incorrectos["cantidad"] >= 3:
            cerrar("PASSWORD_INCORRECTA")
            return

        restantes = 3 - intentos_incorrectos["cantidad"]
        estado_intentos.config(
            text=f"Clave incorrecta. Intentos restantes: {restantes}."
        )
        entrada_clave.focus_set()

    tk.Button(
        botones,
        text="CANCELAR",
        padx=16,
        pady=7,
        command=cerrar,
        **estilo_boton("secundario")
    ).pack(side=tk.RIGHT, padx=(10, 0))

    tk.Button(
        botones,
        text=texto_boton,
        padx=16,
        pady=7,
        command=autenticar,
        **estilo_boton(tipo_boton)
    ).pack(side=tk.RIGHT)

    def al_destruir(event):
        if event.widget == dialogo:
            finalizar_dialogo_administrativo("DESTROY_PROGRAMATICO", dialogo=dialogo)

    dialogo.protocol("WM_DELETE_WINDOW", lambda: cerrar("CANCELADO"))
    dialogo.bind("<Return>", lambda _evento: autenticar())
    dialogo.bind("<Escape>", lambda _evento: cerrar("CANCELADO"))
    dialogo.bind("<Destroy>", al_destruir, add="+")
    vincular_dispatcher_administrativo(dialogo)
    _centrar_dialogo(dialogo, parent, ancho_minimo=520, alto_minimo=260)
    dialogo.grab_set()
    entrada_clave.focus_force()
    registrar_evento_tecnico("ADMIN_DIALOGO_ABIERTO")
    return "break"


def iniciar_contexto_contingencia(
    estado, tipo_operacion, matricula, parent, reanudar=None,
    recuperar_bypass=None,
):
    global contexto_contingencia_proximidad
    contexto_contingencia_proximidad = {
        "estado": estado,
        "operacion": tipo_operacion,
        "matricula": matricula,
        "parent": parent,
        "autorizado": False,
        "reanudado": False,
        "reanudar": reanudar,
        "recuperar_bypass": recuperar_bypass,
        "cerrar_dialogo_fallo": None
    }
    return contexto_contingencia_proximidad


def registrar_dialogo_fallo_contingencia(_dialogo, cerrar_dialogo):
    if contexto_contingencia_proximidad is not None:
        contexto_contingencia_proximidad["cerrar_dialogo_fallo"] = cerrar_dialogo
        contexto_contingencia_proximidad["dialogo"] = _dialogo
        contexto_contingencia_proximidad["ultima_tecla"] = None
        registrar_evento_tecnico(
            "ESTADO_PROXIMIDAD_ACTUAL="
            f"{contexto_contingencia_proximidad['estado']}"
        )
        registrar_evento_tecnico(
            "CONTINGENCIA_ABIERTA_FLAG="
            f"{confirmacion_contingencia_abierta}"
        )
        vincular_dispatcher_administrativo(_dialogo)


def finalizar_contexto_contingencia():
    global contexto_contingencia_proximidad
    autorizado = bool(
        contexto_contingencia_proximidad
        and contexto_contingencia_proximidad["autorizado"]
    )
    contexto_contingencia_proximidad = None
    return autorizado


def es_atajo_administrativo(event):
    """Reconoce Ctrl+Alt+U sin depender de la variante de keysym de Windows."""
    try:
        estado = int(getattr(event, "state", 0))
        return (
            str(getattr(event, "keysym", "")).lower() == "u"
            and bool(estado & 0x0004)  # Control
            # Tk usa Mod1 (0x0008), pero Windows también reporta Alt como 0x20000.
            and bool(estado & (0x0008 | 0x20000))
        )
    except (TypeError, ValueError):
        return False


def es_atajo_salida_administrativa(event):
    """Escape técnico independiente: Ctrl+Alt+Shift+U desde cualquier modal."""
    try:
        estado = int(getattr(event, "state", 0))
        return es_atajo_administrativo(event) and bool(estado & 0x0001)
    except (TypeError, ValueError):
        return False


def registrar_tecla_proximidad(event):
    """Registra metadatos de teclado, nunca el texto introducido por el usuario."""
    contexto = contexto_contingencia_proximidad
    if not contexto:
        return

    huella = (
        getattr(event, "serial", None),
        getattr(event, "time", None),
        getattr(event, "keycode", None),
        getattr(event, "state", None),
        getattr(event, "keysym", None),
    )
    if contexto.get("ultima_tecla") == huella:
        return
    contexto["ultima_tecla"] = huella

    widget = getattr(event, "widget", None)
    try:
        widget_seguro = widget.winfo_class() if widget else "NINGUNO"
    except tk.TclError:
        widget_seguro = "DESTRUIDO"
    registrar_evento_tecnico(
        "TECLA_PROXIMIDAD "
        f"keysym={getattr(event, 'keysym', '')} "
        f"keycode={getattr(event, 'keycode', '')} "
        f"state={getattr(event, 'state', '')} "
        f"widget={widget_seguro}"
    )


def _contexto_visual_administrativo(event=None, origen=None):
    """Resuelve el Toplevel activo sin limitarse al root o a la entrega."""
    candidatos = [origen, getattr(event, "widget", None) if event else None]
    try:
        candidatos.extend((ventana.grab_current(), ventana.focus_get()))
    except tk.TclError:
        pass

    for candidato in candidatos:
        if not _widget_tk_vivo(candidato):
            continue
        try:
            return candidato.winfo_toplevel()
        except tk.TclError:
            continue
    return ventana


def _accion_atajo_administrativo(estado_proximidad):
    """Ctrl+Alt+U siempre pide autorización para omitir proximidad activa."""
    del estado_proximidad
    return "CONTINUAR_SIN_PROXIMIDAD"


def _autorizar_bypass_proximidad(contexto):
    """Autoriza y reanuda la operación que estaba esperando proximidad."""
    if contexto_contingencia_proximidad is not contexto or contexto.get("autorizado"):
        return
    estado = contexto.get("estado")
    registrar_evento_tecnico(
        f"PROXIMIDAD_BYPASS_ADMIN_AUTORIZADO estado={estado}"
    )
    contexto["autorizado"] = True
    registrar_uso_contingencia(
        contexto["operacion"],
        contexto["matricula"],
    )
    cerrar_dialogo_fallo = contexto.get("cerrar_dialogo_fallo")
    if cerrar_dialogo_fallo:
        cerrar_dialogo_fallo()
    reanudar = contexto.get("reanudar")
    if reanudar:
        # No depender de que wait_window vuelva a la pila que abrió el diálogo:
        # el callback administrativo debe reanudar esta operación concreta.
        contexto["reanudado"] = True
        parent = contexto.get("parent")
        try:
            parent.after_idle(reanudar)
        except (AttributeError, tk.TclError):
            try:
                reanudar()
            except Exception as error:
                registrar_evento_tecnico(
                    "DEVOLUCION_BYPASS_CONTINUACION_ERROR "
                    f"tipo={type(error).__name__}"
                )
                recuperar = contexto.get("recuperar_bypass")
                if recuperar:
                    recuperar()


def _solicitar_contingencia_proximidad(contexto, event=None, origen=None):
    estado = contexto.get("estado")
    registrar_evento_tecnico(
        f"CONTINGENCIA_PROXIMIDAD_SOLICITADA estado={estado}"
    )
    dialogo_fallo = contexto.get("dialogo")
    parent = (
        dialogo_fallo
        if _widget_tk_vivo(dialogo_fallo)
        else contexto["parent"]
    )
    return abrir_acceso_administrativo(
        event,
        parent=parent,
        contexto=f"PROXIMIDAD_{estado}",
        accion="CONTINUAR_SIN_PROXIMIDAD",
        al_autorizar=lambda: _autorizar_bypass_proximidad(contexto),
    )


def _solicitar_salida_administrativa(event, origen, contexto):
    parent = _contexto_visual_administrativo(event, origen)
    return abrir_acceso_administrativo(
        event,
        parent=parent,
        contexto=contexto,
        accion="SALIR_SISTEMA",
    )


def manejar_evento_administrativo(
    event,
    origen=None,
    combinacion_confirmada=False,
    forzar_salida=False,
):
    """Entrada común para bindings directos y el respaldo global."""
    registrar_tecla_proximidad(event)
    forzar_salida = forzar_salida or es_atajo_salida_administrativa(event)
    if forzar_salida or combinacion_confirmada or es_atajo_administrativo(event):
        return manejar_atajo_administrativo(
            event,
            origen=origen,
            combinacion_confirmada=combinacion_confirmada,
            forzar_salida=forzar_salida,
        )
    return None


def vincular_dispatcher_administrativo(widget):
    """Da a cada root/Toplevel acceso directo al dispatcher central."""
    if getattr(widget, "_dispatcher_admin_vinculado", False):
        return
    widget._dispatcher_admin_vinculado = True
    widget.bind(
        "<KeyPress>",
        lambda event, origen=widget: manejar_evento_administrativo(event, origen),
        add="+",
    )
    # Estos bindings cubren la traducción nativa de Tk; KeyPress cubre las
    # variantes de state (Alt 0x0008 / 0x20000) observadas en Windows.
    for secuencia in ("<Control-Alt-u>", "<Control-Alt-U>"):
        widget.bind(
            secuencia,
            lambda event, origen=widget: manejar_evento_administrativo(
                event,
                origen,
                combinacion_confirmada=True,
            ),
            add="+",
        )
    for secuencia in ("<Control-Alt-Shift-u>", "<Control-Alt-Shift-U>"):
        widget.bind(
            secuencia,
            lambda event, origen=widget: manejar_evento_administrativo(
                event,
                origen,
                combinacion_confirmada=True,
                forzar_salida=True,
            ),
            add="+",
        )


def manejar_tecla_global(event):
    """Respaldo global; los modales críticos también tienen binding directo."""
    registrar_actividad_usuario(event)
    return manejar_evento_administrativo(event)


def manejar_atajo_administrativo(
    event=None,
    origen=None,
    combinacion_confirmada=False,
    forzar_salida=False,
):
    global atajo_administrativo_en_curso
    global ultimo_atajo_administrativo

    if not combinacion_confirmada and event is not None and not es_atajo_administrativo(event):
        return None

    keysym = str(getattr(event, "keysym", "")) if event else ""
    state = getattr(event, "state", "") if event else ""
    registrar_evento_tecnico(f"TECLA_ADMIN_RECIBIDA keysym={keysym} state={state}")
    registrar_evento_tecnico("ATAJO_CTRL_ALT_U_DETECTADO")
    ahora = time.monotonic()
    if atajo_administrativo_en_curso or ahora - ultimo_atajo_administrativo < 0.20:
        registrar_evento_tecnico("ATAJO_ADMIN_OMITIDO motivo=doble_disparo")
        return "break"

    atajo_administrativo_en_curso = True
    ultimo_atajo_administrativo = ahora
    try:
        contexto_proximidad = contexto_contingencia_proximidad
        if contexto_proximidad and not forzar_salida:
            estado = contexto_proximidad.get("estado")
            registrar_evento_tecnico("ATAJO_ADMIN_DETECTADO contexto=PROXIMIDAD")
            registrar_evento_tecnico(f"ESTADO_PROXIMIDAD_ACTUAL={estado}")
            registrar_evento_tecnico(
                f"CONTINGENCIA_ABIERTA_FLAG={confirmacion_contingencia_abierta}"
            )
            accion = _accion_atajo_administrativo(estado)
            if accion == "CONTINUAR_SIN_PROXIMIDAD":
                return _solicitar_contingencia_proximidad(
                    contexto_proximidad,
                    event,
                    origen,
                )

        if not forzar_salida:
            registrar_evento_tecnico(
                "ATAJO_CTRL_ALT_U_RECHAZADO motivo=SIN_PROXIMIDAD_ACTIVA"
            )
            return "break"

        parent = _contexto_visual_administrativo(event, origen)
        contexto = "LOGIN" if parent == ventana else "MODAL"
        if forzar_salida and contexto_proximidad:
            contexto = f"PROXIMIDAD_{contexto_proximidad.get('estado')}_SALIDA"
        registrar_evento_tecnico(f"ATAJO_ADMIN_DETECTADO contexto={contexto}")
        return _solicitar_salida_administrativa(event, parent, contexto)
    finally:
        atajo_administrativo_en_curso = False


def manejar_atajo_admin(event=None, combinacion_confirmada=False):
    """Alias temporal para integraciones que invocaban el nombre anterior."""
    return manejar_atajo_administrativo(
        event,
        combinacion_confirmada=combinacion_confirmada,
    )


def crear_modal_verificacion_ubicacion(
    parent,
    mensaje="Asegúrate de estar cerca del carrito.",
    titulo="Verificando ubicación del equipo",
    estado="Intento 1 de 5"
):
    dialogo = tk.Toplevel(parent)
    vincular_dispatcher_administrativo(dialogo)
    dialogo.title(titulo)
    dialogo.resizable(False, False)
    dialogo.configure(bg=COLOR_TARJETA)
    dialogo.transient(parent)
    dialogo.attributes("-topmost", True)
    dialogo.protocol("WM_DELETE_WINDOW", lambda: None)
    dialogo.withdraw()

    contenido = tk.Frame(dialogo, bg=COLOR_TARJETA, padx=30, pady=24)
    contenido.pack(fill=tk.BOTH, expand=True)

    if TEMA_INTERFAZ == "moderno":
        indicador = tk.Canvas(
            contenido, width=52, height=52, bg=COLOR_TARJETA,
            highlightthickness=0
        )
        indicador.pack(pady=(0, 12))
        indicador.create_oval(4, 4, 48, 48, fill=COLORES["info_suave"], outline="")
        indicador.create_text(26, 26, text="⌖", fill=COLOR_PRIMARIO,
                              font=("Segoe UI Symbol", 19, "bold"))

    titulo_label = tk.Label(
        contenido,
        text=titulo,
        font=FUENTE_SUBTITULO,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    )
    titulo_label.pack(pady=(0, 14))

    mensaje_label = tk.Label(
        contenido,
        text=mensaje,
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA,
        justify=tk.CENTER,
        wraplength=350
    )
    mensaje_label.pack()

    estado_intento = tk.Label(
        contenido,
        text=estado,
        font=FUENTE_CUERPO_BOLD,
        fg=COLOR_PRIMARIO,
        bg=(COLORES["info_suave"] if TEMA_INTERFAZ == "moderno" else COLOR_TARJETA)
    )
    estado_intento.pack(
        pady=(18, 0),
        ipadx=(16 if TEMA_INTERFAZ == "moderno" else 0),
        ipady=(7 if TEMA_INTERFAZ == "moderno" else 0)
    )

    _centrar_dialogo(
        dialogo,
        parent,
        ancho_minimo=TAMANOS["dialogo_ancho"],
        alto_minimo=(260 if TEMA_INTERFAZ == "moderno" else 210)
    )
    dialogo.grab_set()
    dialogo.focus_force()
    return dialogo, titulo_label, mensaje_label, estado_intento


def actualizar_modal_progreso(dialogo, titulo_label, mensaje_label, estado_label,
                              titulo, mensaje, estado):
    dialogo.title(titulo)
    titulo_label.config(text=titulo)
    mensaje_label.config(text=mensaje)
    estado_label.config(text=estado)


def convertir_modal_en_confirmacion(
    dialogo,
    titulo_label,
    mensaje_label,
    estado_label,
    nombre,
    al_confirmar,
    al_cancelar
):
    activar_timeout_inactividad()
    registrar_ventana_flujo_alumno(dialogo)
    dialogo.title("Confirmación")
    titulo_label.config(text="Confirmación")
    mensaje_label.config(text=f"¿Eres {nombre}?")
    estado_label.pack_forget()

    botones = tk.Frame(estado_label.master, bg=COLOR_TARJETA)
    botones.pack(fill=tk.X, pady=(22, 0))
    respondido = {"valor": False}

    def responder(confirmado):
        if respondido["valor"]:
            return
        respondido["valor"] = True
        cerrar_modal_verificacion_ubicacion(dialogo)
        if confirmado:
            al_confirmar()
        else:
            al_cancelar()

    tk.Button(
        botones,
        text="NO",
        padx=26,
        pady=8,
        command=lambda: responder(False),
        **estilo_boton("secundario")
    ).pack(side=tk.RIGHT, padx=(12, 0))

    tk.Button(
        botones,
        text="SÍ",
        padx=26,
        pady=8,
        command=lambda: responder(True),
        **estilo_boton("principal")
    ).pack(side=tk.RIGHT)

    dialogo.protocol("WM_DELETE_WINDOW", lambda: responder(False))
    dialogo.bind("<Return>", lambda _evento: responder(True))
    dialogo.bind("<Escape>", lambda _evento: responder(False))
    _centrar_dialogo(dialogo, dialogo.master, ancho_minimo=430, alto_minimo=210)
    dialogo.focus_force()


def cerrar_modal_verificacion_ubicacion(dialogo):
    try:
        dialogo.grab_release()
    except tk.TclError:
        pass
    try:
        dialogo.destroy()
    except tk.TclError:
        pass


def centrar_ventana(ventana, ancho, alto):
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


# --- FUNCIONES BASE (SIN CAMBIOS) ---
def apagar_windows():
    if threading.current_thread() is not threading.main_thread():
        ventana.after(0, apagar_windows)
        return

    if MODO_PRUEBA:
        registrar_evento_tecnico("APAGADO_SIMULADO modo_prueba=True")
        mostrar_exito(
            "Modo de prueba",
            "La devolución terminó correctamente. En funcionamiento normal, "
            "la laptop se apagaría en este momento."
        )
        # Reinicia sólo el flujo de la aplicación para poder ejecutar otra
        # prueba; nunca cierra sesión de Windows ni toca el proceso del sistema.
        crear_pantalla_login()
        return

    os.system("shutdown /s /t 3")


def contar_usos_alumno(matricula):
    """
    Cuenta cuántas veces el alumno ha usado el sistema
    con base en la hoja 'Registros'.
    """
    try:
        if hoja_registros is None:
            return 0

        matriculas = hoja_registros.col_values(1)
        return matriculas.count(matricula)

    except Exception as e:
        print(f"Error al contar usos del alumno: {e}")
        return 0


def _calcular_alto_aviso(alto_requerido, alto_pantalla):
    return min(max(alto_requerido + 16, 230), alto_pantalla - 80)


def mostrar_ventana_control_unificada(
    matricula,
    nombre,
    tiene_pendiente,
    fecha_entrada,
    laptop_id,
    no_entregas,
    estado
):
    ventana_ctrl = tk.Toplevel(ventana)
    vincular_dispatcher_administrativo(ventana_ctrl)
    registrar_ventana_flujo_alumno(ventana_ctrl)
    ventana_ctrl.title("Aviso importante")
    ventana_ctrl.resizable(False, False)
    ventana_ctrl.configure(bg="#ffffff")
    
    # Aplicar estilo moderno
    ventana_ctrl.configure(bg=COLOR_FONDO)
    
    # Hacerla modal respecto a la ventana principal
    ventana_ctrl.transient(ventana)
    ventana_ctrl.grab_set()

    ancho = 500
    frame = tk.Frame(ventana_ctrl, bg=COLOR_TARJETA)
    frame.pack(fill=tk.BOTH, expand=True)

    contenido = tk.Frame(frame, bg=COLOR_TARJETA, padx=30, pady=12)
    contenido.pack(fill=tk.X)

    boton_frame = tk.Frame(frame, bg=COLOR_TARJETA, padx=40, pady=5)
    boton_frame.pack(side=tk.BOTTOM, fill=tk.X)


    # USUARIO SANCIONADO
    if estado == "SANCIONADO":
        titulo = "USUARIO SANCIONADO"
        color = COLOR_ERROR
        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n\n"
            "Has excedido el número permitido.\n\n"
            "Acude con el administrador."
        )
    else:
        # AVISO GENERAL
        titulo = "AVISO IMPORTANTE"
        color = COLOR_ADVERTENCIA

        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n"
            f"Estado: {estado}\n\n"
            "Recuerda entregar correctamente la laptop al finalizar su uso.\n\n"
            "Al alcanzar el límite permitido, el acceso será bloqueado."
        )


        if tiene_pendiente:
            mensaje = (
                "Entrega pendiente detectada\n\n"
                f"Laptop: {laptop_id}\n"
                f"Fecha: {formatear_fecha(fecha_entrada)}\n\n"
            ) + mensaje

    # Icono

    tk.Label(
        contenido,
        text=titulo,
        font=("Segoe UI", 16, "bold"),
        fg=color,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 10))

    tk.Label(
        contenido,
        text=mensaje,
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA,
        justify=tk.CENTER,
        wraplength=420
    ).pack(pady=(8, 6), fill=tk.X)


    
    tk.Button(
        boton_frame,
        text="ENTENDIDO",
        font=("Segoe UI", 13, "bold"),   # ? más equilibrado
        bg=COLOR_PRIMARIO,
        fg="white",
        bd=0,
        cursor="hand2",
        activebackground=COLOR_SECUNDARIO,
        activeforeground="white",
        relief="flat",
        command=ventana_ctrl.destroy
    ).pack(
        fill=tk.X,
        padx=90,        # ? más angosto
        pady=(8, 5),
        ipady=6         # ? altura fina y controlada
    )





    ventana_ctrl.update_idletasks()
    pantalla_ancho = ventana_ctrl.winfo_screenwidth()
    pantalla_alto = ventana_ctrl.winfo_screenheight()
    alto = _calcular_alto_aviso(ventana_ctrl.winfo_reqheight(), pantalla_alto)
    x = (pantalla_ancho - ancho) // 2
    y = (pantalla_alto - alto) // 2
    ventana_ctrl.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana_ctrl.focus_force()

    # ESPERAR a que el usuario cierre la ventana
    ventana_ctrl.wait_window()

def matricula_exenta_no_entrega(matricula):
    return str(matricula).strip() in MATRICULAS_EXENTAS_NO_ENTREGA


def incrementar_no_entregas(matricula, contador_objetivo):
    """Obsoleto: los incrementos se aplican sólo en Apps Script."""
    raise RuntimeError("NO_ENTREGA_DEBE_USAR_COORDINADOR")
    
def _contador_no_entregas_actual(matricula):
    matriculas = hoja_alumnos.col_values(1)
    if matricula not in matriculas:
        raise ErrorConsultaGoogle("La matrícula no existe en la hoja Alumnos")
    fila = matriculas.index(matricula) + 1
    valor = hoja_alumnos.cell(fila, 5).value
    return int(valor) if valor else 0


def _objetivo_estado_intermedio(observacion, prefijo):
    if not observacion.startswith(prefijo + ":"):
        return None
    try:
        return int(observacion.split(":", 1)[1])
    except (TypeError, ValueError):
        raise ErrorConsultaGoogle("Estado intermedio de no entrega inválido")


def _finalizar_observacion(numero_fila, observacion_final):
    raise RuntimeError("OBSERVACION_DEBE_USAR_COORDINADOR")


def cerrar_sesion_anterior_y_contar_no_entrega(matricula):
    with bloqueo_local_registro_prestamo():
        return _cerrar_sesion_anterior_y_contar_no_entrega(matricula)


def _cerrar_sesion_anterior_y_contar_no_entrega(matricula):
    """
    Cierra la sesión activa anterior del alumno,
    registra NO ENTREGA y deja evidencia.
    """
    try:
        if hoja_registros is None or hoja_alumnos is None:
            raise ErrorConsultaGoogle("Las hojas de Google no están disponibles")

        registros = hoja_registros.get_all_values()
        hora_actual, _ = obtener_hora_internet()

        # Las escrituras críticas se delegan al coordinador; gspread sólo leyó la sesión.
        for fila in reversed(registros):
            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula and fila[COL_HORA_SALIDA].strip() == "" and observacion_es_normal(fila[COL_OBSERVACION]):
                datos = {"matricula": matricula, "nombre": fila[COL_NOMBRE], "fecha": fila[COL_FECHA], "hora_ingreso": fila[COL_HORA_INGRESO], "hora_salida": hora_actual, "laptop_id": fila[COL_LAPTOP_ID], "observacion": "CIERRE_AUTOMATICO_POR_NUEVA_SESION"}
                clave = f"{matricula}:{fila[COL_LAPTOP_ID]}:{fila[COL_FECHA]}:{fila[COL_HORA_INGRESO]}"
                if not matricula_exenta_no_entrega(matricula):
                    crear_y_enviar_evento_critico("NO_ENTREGA", matricula, fila[COL_NOMBRE], fila[COL_FECHA], hora_actual, fila[COL_LAPTOP_ID], datos, clave)
                crear_y_enviar_evento_critico("CIERRE_AUTOMATICO", matricula, fila[COL_NOMBRE], fila[COL_FECHA], hora_actual, fila[COL_LAPTOP_ID], datos, clave)
                return True
        return False

        for i in reversed(range(len(registros))):
            fila = registros[i]

            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                observacion = fila[COL_OBSERVACION].strip()
                if observacion == "CIERRE_AUTOMATICO_POR_NUEVA_SESION":
                    return True

                objetivo_pendiente = _objetivo_estado_intermedio(
                    observacion, "PROCESANDO_CIERRE_AUTOMATICO"
                )
                if objetivo_pendiente is not None:
                    if not matricula_exenta_no_entrega(matricula):
                        incrementar_no_entregas(matricula, objetivo_pendiente)
                    _finalizar_observacion(
                        i + 1, "CIERRE_AUTOMATICO_POR_NUEVA_SESION"
                    )
                    return True

                if (
                    fila[COL_HORA_SALIDA].strip() == ""
                    and observacion_es_normal(fila[COL_OBSERVACION])
                ):
                    numero_fila = i + 1
                    objetivo = _contador_no_entregas_actual(matricula) + 1
                    estado_actual = hoja_registros.get(
                        f"E{numero_fila}:G{numero_fila}"
                    )
                    if not estado_actual or len(estado_actual[0]) <= 2:
                        raise ErrorConsultaGoogle("No se pudo releer la sesión anterior")
                    salida_actual = estado_actual[0][0].strip()
                    observacion_actual = estado_actual[0][2].strip()
                    if observacion_actual == "CIERRE_AUTOMATICO_POR_NUEVA_SESION":
                        return True
                    objetivo_existente = _objetivo_estado_intermedio(
                        observacion_actual, "PROCESANDO_CIERRE_AUTOMATICO"
                    )
                    if objetivo_existente is not None:
                        if not matricula_exenta_no_entrega(matricula):
                            incrementar_no_entregas(matricula, objetivo_existente)
                        _finalizar_observacion(
                            numero_fila, "CIERRE_AUTOMATICO_POR_NUEVA_SESION"
                        )
                        return True
                    if salida_actual or not observacion_es_normal(observacion_actual):
                        return False

                    estado_intermedio = f"PROCESANDO_CIERRE_AUTOMATICO:{objetivo}"
                    raise RuntimeError("CIERRE_AUTOMATICO_DEBE_USAR_COORDINADOR")
                    verificacion = hoja_registros.get(f"E{numero_fila}:G{numero_fila}")
                    if (
                        not verificacion
                        or len(verificacion[0]) <= 2
                        or verificacion[0][0].strip() != hora_actual
                        or verificacion[0][2].strip() not in (
                            estado_intermedio,
                            "CIERRE_AUTOMATICO_POR_NUEVA_SESION",
                        )
                    ):
                        raise ErrorConsultaGoogle("No se pudo verificar el cierre automático")

                    if verificacion[0][2].strip() == "CIERRE_AUTOMATICO_POR_NUEVA_SESION":
                        return True
                    if not matricula_exenta_no_entrega(matricula):
                        incrementar_no_entregas(matricula, objetivo)
                    _finalizar_observacion(
                        numero_fila, "CIERRE_AUTOMATICO_POR_NUEVA_SESION"
                    )
                    return True
                break

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al cerrar sesión anterior: {e}")
        raise ErrorConsultaGoogle("No se pudo cerrar la sesión anterior") from e

    return False


def obtener_control_alumno(matricula):
    """
    Obtiene el número de no entregas y el estado del alumno
    desde la hoja 'Alumnos'.
    """
    try:
        if hoja_alumnos is None:
            raise ErrorConsultaGoogle("La hoja Alumnos no está disponible")

        matriculas = hoja_alumnos.col_values(1)
        if matricula in matriculas:
            fila = matriculas.index(matricula) + 1

            no_entregas = hoja_alumnos.cell(fila, 5).value
            estado = hoja_alumnos.cell(fila, 7).value

            no_entregas = int(no_entregas) if no_entregas else 0
            estado = normalizar_estado_alumno(estado, lectura_exitosa=True)
            actualizar_cache_persona(
                matricula, no_entregas=no_entregas, estado=estado
            )

            return no_entregas, estado

        raise ErrorConsultaGoogle("La matrícula no existe en la hoja Alumnos")

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al obtener control del alumno: {e}")
        raise ErrorConsultaGoogle("No se pudo consultar el control del alumno") from e

def mostrar_instrucciones_iniciales(matricula=""):
    ventana_info = tk.Toplevel(ventana)
    vincular_dispatcher_administrativo(ventana_info)
    ventana_info.title("Instrucciones de Uso")
    ventana_info.resizable(False, False)
    ventana_info.configure(bg=COLOR_FONDO)

    # Modal correcta (SIN topmost)
    ventana_info.transient(ventana)
    ventana_info.grab_set()
    ventana_info.focus_force()

    ancho = 560
    alto = 400

    x = (ventana_info.winfo_screenwidth() - ancho) // 2
    y = (ventana_info.winfo_screenheight() - alto) // 2
    ventana_info.geometry(f"{ancho}x{alto}+{x}+{y}")

    frame = tk.Frame(
        ventana_info,
        bg=COLOR_TARJETA,
        padx=35,
        pady=30
    )
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        frame,
        text="INSTRUCCIONES IMPORTANTES",
        font=("Segoe UI", 16, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 7))

    texto = (
    "• La laptop es responsabilidad del alumno durante su uso.\n"
    "• Cualquier daño o falla debe reportarse de inmediato.\n"
    "• NO cierres el sistema ni apagues la computadora manualmente.\n"
    "• Siempre usa el botón “ENTREGAR Y APAGAR”.\n"
    "• No entregar la laptop genera NO ENTREGA y posibles sanciones.\n"
    )

    texto_widget = tk.Text(
    frame,
    font=FUENTE_CUERPO,
    fg=COLOR_TEXTO_SECUNDARIO,
    bg=COLOR_TARJETA,
    wrap=tk.WORD,
    relief=tk.FLAT,
    height=8,        # clave
    padx=20,
    pady=4
)
    texto_widget.pack(pady=6, fill=tk.X)

    texto_widget.insert(tk.END, texto)
    texto_widget.config(state=tk.DISABLED)




    tk.Button(
        frame,
        text="ENTENDIDO",
        font=FUENTE_BOTON,
        bg=COLOR_EXITO,
        fg="white",
        activebackground="#2ecc71",
        padx=35,
        pady=12,
        bd=0,
        cursor="hand2",
        command=ventana_info.destroy
    ).pack(pady=20)

def bloquear_alt_f4(event):
    """
    Bloquea la combinación Alt + F4 para usuarios normales.
    """
    return "break"

def cargar_logo(ruta_imagen, ancho, alto):
    """Carga y redimensiona el logo"""
    if not PIL_DISPONIBLE:
        return None
    try:
        imagen = Image.open(ruta_imagen)
        imagen = imagen.resize((ancho, alto), Image.LANCZOS)
        return ImageTk.PhotoImage(imagen)
    except Exception as e:
        print(f"Error cargando logo: {e}")
        return None


def crear_fondo_marca_agua_login(ruta_imagen, ancho, alto):
    """Genera una marca de agua ligera para el fondo exterior del login."""
    if not PIL_DISPONIBLE or not FONDO_MARCA_AGUA_LOGIN:
        return None

    try:
        fondo = Image.new("RGB", (ancho, alto), COLOR_FONDO)
        logo = Image.open(ruta_imagen).convert("RGBA")

        mascara_recorte = ImageChops.multiply(
            ImageOps.invert(ImageOps.grayscale(logo)),
            logo.getchannel("A")
        )
        limites_logo = mascara_recorte.getbbox()
        if limites_logo:
            logo = logo.crop(limites_logo)

        logo.thumbnail((72, 58), Image.LANCZOS)

        escala_grises = ImageOps.grayscale(logo)
        mascara_visible = ImageChops.multiply(
            ImageOps.invert(escala_grises),
            logo.getchannel("A")
        )
        mascara = mascara_visible.point(
            lambda valor: int(valor * 0.17)
        )
        marca = Image.new("RGBA", logo.size, (151, 180, 201, 0))
        marca.putalpha(mascara)
        marca = marca.rotate(-18, expand=True, resample=Image.BICUBIC)

        paso_x, paso_y = 145, 118
        for fila, y in enumerate(range(-20, alto + paso_y, paso_y)):
            desplazamiento = 70 if fila % 2 else 0
            for x in range(-70 + desplazamiento, ancho + paso_x, paso_x):
                fondo.paste(marca, (x, y), marca)

        return ImageTk.PhotoImage(fondo)
    except Exception as e:
        print(f"Error creando marca de agua del login: {e}")
        return None

def verificar_internet():
    urls_prueba = [
        "https://www.google.com",
        "https://sheets.googleapis.com",
        "https://www.microsoft.com"
    ]

    for url in urls_prueba:
        try:
            respuesta = requests.get(
                url,
                timeout=2,
                verify=certifi.where(),
                headers={"User-Agent": "Mozilla/5.0"},
                allow_redirects=True
            )
            if respuesta.status_code < 500:
                actualizar_diagnostico_conectividad("internet", INTERNET_OK)
                return True
                
        except Exception as e:
            pass

    actualizar_diagnostico_conectividad("internet", INTERNET_ERROR)
    return False


def registrar_log_actualizacion(mensaje):
    marca_tiempo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ruta_log = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        ARCHIVO_LOG_ACTUALIZACION
    )
    try:
        with open(ruta_log, "a", encoding="utf-8") as archivo:
            archivo.write(f"[{marca_tiempo}] [ACTUALIZACION] {mensaje}\n")
    except OSError:
        pass


def ejecutar_git_segundo_plano(argumentos, timeout=15):
    directorio = os.path.dirname(os.path.abspath(__file__))
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
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        check=False
    )


def comprobar_actualizacion_segundo_plano():
    if not actualizacion_lock.acquire(blocking=False):
        return

    try:
        registrar_log_actualizacion("comprobando...")
        try:
            respuesta = requests.get(
                "https://github.com",
                timeout=3,
                verify=certifi.where(),
                headers={"User-Agent": "RegistroLaptop-Updater"}
            )
            if respuesta.status_code >= 500:
                registrar_log_actualizacion("sin Internet")
                return
        except requests.RequestException:
            registrar_log_actualizacion("sin Internet")
            return

        resultado_fetch = ejecutar_git_segundo_plano(
            ["fetch", "origin", "main", "--quiet"],
            timeout=15
        )
        if resultado_fetch.returncode != 0:
            detalle = resultado_fetch.stderr or resultado_fetch.stdout
            detalle = " ".join((detalle or "error desconocido").split())[:300]
            registrar_log_actualizacion(f"error en fetch: {detalle}")
            return

        pendientes = ejecutar_git_segundo_plano(
            ["rev-list", "--count", "HEAD..origin/main"],
            timeout=5
        )
        if pendientes.returncode != 0:
            registrar_log_actualizacion("error comprobando la versión remota")
            return

        try:
            cantidad_pendiente = int(pendientes.stdout.strip() or "0")
        except ValueError:
            registrar_log_actualizacion("error interpretando la versión remota")
            return

        if cantidad_pendiente == 0:
            registrar_log_actualizacion("sin cambios")
            return

        directorio = os.path.dirname(os.path.abspath(__file__))
        ruta_marker = os.path.join(directorio, ".actualizacion_pendiente")
        ruta_temporal = ruta_marker + ".tmp"
        with open(ruta_temporal, "w", encoding="ascii") as archivo:
            archivo.write("pendiente\n")
        os.replace(ruta_temporal, ruta_marker)
        registrar_log_actualizacion(
            "nueva versión detectada; actualización marcada como pendiente"
        )
    except FileNotFoundError:
        registrar_log_actualizacion("error: Git no está disponible")
    except subprocess.TimeoutExpired:
        registrar_log_actualizacion("error: tiempo agotado al comprobar")
    except Exception as error:
        registrar_log_actualizacion(f"error inesperado: {error}")
    finally:
        actualizacion_lock.release()


def iniciar_actualizacion_en_segundo_plano():
    threading.Thread(
        target=comprobar_actualizacion_segundo_plano,
        daemon=True
    ).start()


def conectar_google_sheets(internet_verificado=False):
    global hoja_alumnos, hoja_registros, hoja_configuracion_sistema

    if not internet_verificado and not verificar_internet():
        cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        return False

    with conexion_google_lock:
        if hoja_alumnos is not None and hoja_registros is not None:
            if verificar_conexion_base_datos():
                return True
            hoja_alumnos = None
            hoja_registros = None

        cambiar_estado("Verificando conexión...", COLOR_ADVERTENCIA)
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                RUTA_CREDENCIALES, scope
            )

            client = gspread.authorize(creds)
            client.http_client.set_timeout((3, 10))

            sheet = client.open("Control de Laptops")

            hoja_alumnos = sheet.worksheet("Alumnos")
            hoja_registros = sheet.worksheet("Registros")
            try:
                hoja_configuracion_sistema = sheet.worksheet("Configuracion_Sistema")
            except Exception:
                # El cliente jamás crea esta hoja ni inventa credenciales.
                hoja_configuracion_sistema = None

            actualizar_diagnostico_conectividad("sheets", SHEETS_OK)
            actualizar_configuracion_sincronizacion()
            iniciar_sincronizacion_pendientes()
            cambiar_estado("Conectado", COLOR_EXITO)
            return True
        except FileNotFoundError as error:
            hoja_alumnos = None
            hoja_registros = None
            hoja_configuracion_sistema = None
            actualizar_diagnostico_conectividad("sheets", SHEETS_AUTH_ERROR)
            registrar_evento_tecnico(
                f"SHEETS_CREDENCIALES_ERROR={type(error).__name__}: {error}"
            )
            cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
            return False
        except Exception as e:
            hoja_alumnos = None
            hoja_registros = None
            hoja_configuracion_sistema = None
            actualizar_diagnostico_conectividad(
                "sheets", clasificar_error_sheets(e)
            )
            cambiar_estado("Error en la conexión", COLOR_ERROR)
            return False


def verificar_conexion_base_datos():
    """Comprueba mediante una lectura remota que Google Sheets responde."""
    if hoja_alumnos is None or hoja_registros is None:
        return False
    try:
        hoja_alumnos.acell("A1")
        actualizar_diagnostico_conectividad("sheets", SHEETS_OK)
        return True
    except Exception as error:
        actualizar_diagnostico_conectividad(
            "sheets", clasificar_error_sheets(error)
        )
        print(f"Error al verificar Google Sheets: {error}")
        return False


def verificar_sheets_con_reintentos(max_intentos=3, callback_intento=None):
    """Verifica Sheets remotamente con pausas breves, sin realizar escrituras."""
    pausas = (0, 1, 2)
    for intento in range(1, max_intentos + 1):
        if intento > 1:
            time.sleep(pausas[min(intento - 1, len(pausas) - 1)])
        if callback_intento:
            callback_intento(intento, max_intentos)

        conectado = verificar_conexion_base_datos()
        if not conectado:
            conectado = conectar_google_sheets(internet_verificado=True)
            if conectado:
                conectado = verificar_conexion_base_datos()

        registrar_evento_tecnico(
            f"SHEETS_CONEXION intento={intento} resultado={'OK' if conectado else 'ERROR'}"
        )
        if conectado:
            return True
    return False

def verificar_conexion_periodicamente():
    global verificar_conexion_activo, aviso_internet
    
    while verificar_conexion_activo:
        if verificar_internet():
            if hoja_alumnos is None or hoja_registros is None:
                if conectar_google_sheets(internet_verificado=True):
                    # Cerrar aviso de internet si está abierto
                    ejecutar_en_ui(cerrar_aviso_internet)
                    
            elif verificar_conexion_base_datos():
                cambiar_estado("Conectado", COLOR_EXITO)
            else:
                cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        else:
            cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        
        # Una lectura de Sheets cada dos segundos agota rápidamente la cuota y
        # alterna falsos SHEETS_API_ERROR/429. La cola inmediata cubre eventos
        # nuevos; esta verificación sólo es salud de respaldo.
        time.sleep(60)

def detener_verificacion_conexion():
    global verificar_conexion_activo
    verificar_conexion_activo = False


def cerrar_aplicacion():
    """Finaliza Tkinter y permite que termine el proceso principal."""
    global cierre_aplicacion_en_curso
    if cierre_aplicacion_en_curso:
        return

    cierre_aplicacion_en_curso = True
    registrar_evento_tecnico("CIERRE_APLICACION_INICIADO")
    detener_verificacion_conexion()
    cancelar_timeout_inactividad()
    try:
        if ventana.winfo_exists():
            ventana.destroy()
    except (tk.TclError, NameError):
        pass
    registrar_evento_tecnico("PROCESO_FINALIZANDO")

def ejecutar_en_ui(func, *args):
    try:
        if threading.current_thread() is threading.main_thread():
            func(*args)
        else:
            ventana.after(0, lambda: func(*args))
    except:
        pass


def registrar_ventana_flujo_alumno(dialogo):
    if not timeout_inactividad_activo:
        return
    ventanas_flujo_alumno.add(dialogo)

    def olvidar_dialogo(event):
        if event.widget == dialogo:
            ventanas_flujo_alumno.discard(dialogo)

    dialogo.bind("<Destroy>", olvidar_dialogo, add="+")


def cancelar_timeout_inactividad():
    global timeout_inactividad_id, timeout_inactividad_activo
    timeout_inactividad_activo = False
    if timeout_inactividad_id is not None:
        try:
            ventana.after_cancel(timeout_inactividad_id)
        except (tk.TclError, ValueError):
            pass
        timeout_inactividad_id = None


def reiniciar_timeout_inactividad(event=None):
    global timeout_inactividad_id
    if not timeout_inactividad_activo:
        return
    if timeout_inactividad_id is not None:
        try:
            ventana.after_cancel(timeout_inactividad_id)
        except (tk.TclError, ValueError):
            pass
    timeout_inactividad_id = ventana.after(
        TIEMPO_INACTIVIDAD_MS,
        procesar_timeout_inactividad
    )


def activar_timeout_inactividad():
    global timeout_inactividad_activo, flujo_cancelado_por_inactividad
    timeout_inactividad_activo = True
    flujo_cancelado_por_inactividad = False
    reiniciar_timeout_inactividad()


def procesar_timeout_inactividad():
    global timeout_inactividad_id, timeout_inactividad_activo
    global flujo_cancelado_por_inactividad, procesando_sesion
    global contexto_contingencia_proximidad, confirmacion_contingencia_abierta

    if not timeout_inactividad_activo:
        return

    timeout_inactividad_id = None
    timeout_inactividad_activo = False
    flujo_cancelado_por_inactividad = True
    contexto_contingencia_proximidad = None
    confirmacion_contingencia_abierta = False
    procesando_sesion = False

    # El admin puede ser hijo de uno de estos modales. Límpialo antes de que
    # Tk lo destruya indirectamente junto con su padre.
    finalizar_dialogo_administrativo("TIMEOUT", restaurar=False)

    for dialogo in list(ventanas_flujo_alumno):
        try:
            if dialogo.winfo_exists():
                dialogo.destroy()
        except tk.TclError:
            pass
    ventanas_flujo_alumno.clear()

    try:
        entrada.delete(0, tk.END)
        acepta_estado_equipo.set(False)
        btn_entrar.config(state="normal", text="INGRESAR AL SISTEMA")
        cambiar_estado(
            "Proceso cancelado por inactividad",
            COLOR_ADVERTENCIA
        )
        ventana.deiconify()
        ventana.lift()
        ventana.after(100, entrada.focus_set)
    except (tk.TclError, NameError):
        pass


def registrar_actividad_usuario(event=None):
    reiniciar_timeout_inactividad()


def cerrar_aviso_internet():
    global aviso_internet
    try:
        if aviso_internet and aviso_internet.winfo_exists():
            aviso_internet.destroy()
        aviso_internet = None
    except:
        aviso_internet = None
def obtener_hora_internet():
    try:
        response = requests.get(
            "http://worldtimeapi.org/api/timezone/America/Chihuahua",
            timeout=5,
            verify=certifi.where(),
            headers={"User-Agent": "Mozilla/5.0"}
        )
        hora_utc = datetime.fromisoformat(response.json()["datetime"].split(".")[0])
        return zona_horaria.localize(hora_utc).strftime("%H:%M:%S"), zona_horaria.localize(hora_utc).strftime("%Y-%m-%d")
    except:
        ahora = datetime.now(zona_horaria)
        return ahora.strftime("%H:%M:%S"), ahora.strftime("%Y-%m-%d")

def buscar_nombre(matricula):
    try:
        if hoja_alumnos is None:
            return None
        celdas = hoja_alumnos.col_values(1)
        if matricula in celdas:
            row = celdas.index(matricula) + 1
            nombre = hoja_alumnos.cell(row, 2).value
            actualizar_cache_persona(matricula, nombre=nombre)
            return nombre
    except Exception as e:
        print(f"Error al buscar nombre: {e}")
        return None
    return None

def buscar_rol(matricula):
    """
    Obtiene el rol del usuario desde Google Sheets
    Columna C
    """
    try:
        if hoja_alumnos is None:
            return "ALUMNO"

        celdas = hoja_alumnos.col_values(1)
        if matricula in celdas:
            row = celdas.index(matricula) + 1
            rol = hoja_alumnos.cell(row, 3).value
            rol = rol.strip().upper() if rol else "ALUMNO"
            actualizar_cache_persona(matricula, rol=rol)
            return rol
    except Exception as e:
        print(f"Error al buscar rol: {e}")
        return "ALUMNO"


def consultar_datos_identidad(matricula):
    """Lee el rol y, solo para alumnos, la CURP desde la misma fila."""
    try:
        if hoja_alumnos is None:
            raise ErrorConsultaGoogle("La hoja Alumnos no está disponible")

        matriculas = hoja_alumnos.col_values(1)
        if matricula not in matriculas:
            return None, None

        fila = matriculas.index(matricula) + 1
        rol = normalizar_rol(hoja_alumnos.cell(fila, 3).value)
        curp = hoja_alumnos.cell(fila, 4).value if rol_requiere_curp(rol) else None
        # Sólo se persisten los dos caracteres usados en la validación, no la CURP.
        curp_ultimos_2 = str(curp).strip()[-2:] if curp else None
        actualizar_cache_persona(
            matricula, rol=rol, curp_ultimos_2=curp_ultimos_2
        )
        return rol, curp
    except ErrorConsultaGoogle:
        raise
    except Exception as error:
        raise ErrorConsultaGoogle("No se pudo consultar la identidad") from error


def buscar_curp(matricula):
    """
    Obtiene la CURP completa desde Google Sheets
    Columna D
    """
    try:
        if hoja_alumnos is None:
            return None

        celdas = hoja_alumnos.col_values(1)
        if matricula in celdas:
            row = celdas.index(matricula) + 1
            return hoja_alumnos.cell(row, 4).value
    except Exception as e:
        print(f"Error al buscar CURP: {e}")
        return None


def validar_curp_ultimos_2(curp_real, curp_ingresada):
    if not curp_real or len(curp_real) < 2:
        return False
    return curp_real[-2:] == curp_ingresada.strip()


def pedir_curp_ultimos_2(parent):
    activar_timeout_inactividad()
    ventana_curp = tk.Toplevel(parent)
    vincular_dispatcher_administrativo(ventana_curp)
    registrar_ventana_flujo_alumno(ventana_curp)
    ventana_curp.title("Validación de identidad")
    ventana_curp.resizable(False, False)
    ventana_curp.configure(bg=COLOR_FONDO)
    ventana_curp.transient(parent)
    ventana_curp.attributes("-topmost", True)
    ventana_curp.withdraw()

    tarjeta = tk.Frame(
        ventana_curp,
        bg=COLOR_TARJETA,
        highlightthickness=1,
        highlightbackground=COLOR_BORDE,
        padx=34,
        pady=28
    )
    tarjeta.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    tk.Label(
        tarjeta,
        text="Validación de identidad",
        font=("Segoe UI", 17, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 8))

    tk.Label(
        tarjeta,
        text="Por seguridad, ingresa los últimos 2 caracteres de tu CURP.",
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA,
        justify=tk.CENTER,
        wraplength=360
    ).pack(pady=(0, 22))

    tk.Label(
        tarjeta,
        text="ÚLTIMOS 2 CARACTERES",
        font=FUENTE_PEQ_BOLD,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(anchor="w")

    resultado = {"valor": None}

    # VARIABLE CONTROLADA
    curp_var = tk.StringVar()

    def controlar_texto(*args):
        valor = curp_var.get().upper()

        # Limitar a 2 caracteres
        if len(valor) > 2:
            valor = valor[:2]

        # Evitar loop infinito
        if curp_var.get() != valor:
            curp_var.set(valor)

    # Vincular control
    curp_var.trace_add("write", controlar_texto)

    estilo_campo_curp = estilo_entrada()
    estilo_campo_curp["font"] = ("Segoe UI", 18, "bold")
    entrada = tk.Entry(
        tarjeta,
        textvariable=curp_var,
        justify="center",
        **estilo_campo_curp
    )
    entrada.pack(fill=tk.X, pady=(7, 20), ipady=10)

    def confirmar(event=None):
        valor = curp_var.get().strip()

        if len(valor) != 2:
            mostrar_error(
                "Dato inválido",
                "Debes ingresar exactamente 2 caracteres.",
                parent=ventana_curp
            )
            entrada.focus_force()
            entrada.selection_range(0, tk.END)
            return

        resultado["valor"] = valor
        ventana_curp.attributes("-topmost", False)
        ventana_curp.destroy()

    entrada.bind("<Return>", confirmar)

    tk.Button(
        tarjeta,
        text="VALIDAR",
        padx=25,
        pady=10,
        command=confirmar,
        **estilo_boton("principal")
    ).pack(fill=tk.X)

    _centrar_dialogo(ventana_curp, parent, ancho_minimo=440, alto_minimo=310)
    ventana_curp.grab_set()
    ventana_curp.lift()

    def enfocar_entrada():
        entrada.focus_force()
        entrada.selection_range(0, tk.END)
        entrada.icursor(tk.END)

    ventana_curp.after_idle(enfocar_entrada)
    ventana_curp.wait_window()
    return resultado["valor"]

def verificar_entrega_pendiente(matricula):
    try:
        if hoja_registros is None:
            raise ErrorConsultaGoogle("La hoja Registros no está disponible")

        registros = hoja_registros.get_all_values()

        for i in reversed(range(len(registros))):
            fila = registros[i]
            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                observacion = fila[COL_OBSERVACION].strip()
                if hora_salida == "" and observacion_es_normal(observacion):
                    fecha_entrada = fila[COL_FECHA]
                    laptop_id = fila[COL_LAPTOP_ID]
                    return True, fecha_entrada, laptop_id

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al verificar entrega pendiente: {e}")
        raise ErrorConsultaGoogle("No se pudo verificar la entrega pendiente") from e

    return False, None, None


def procesar_no_entrega_si_corresponde(matricula):
    with bloqueo_local_registro_prestamo():
        return _procesar_no_entrega_si_corresponde(matricula)


def _procesar_no_entrega_si_corresponde(matricula):
    try:
        if hoja_registros is None or hoja_alumnos is None:
            raise ErrorConsultaGoogle("Las hojas de Google no están disponibles")

        registros = hoja_registros.get_all_values()

        for fila in reversed(registros):
            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula and fila[COL_HORA_SALIDA].strip() == "" and observacion_es_normal(fila[COL_OBSERVACION]) and not matricula_exenta_no_entrega(matricula):
                fecha, hora = _fecha_hora_local()
                datos = {"matricula": matricula, "nombre": fila[COL_NOMBRE], "fecha": fila[COL_FECHA], "hora_ingreso": fila[COL_HORA_INGRESO], "hora_salida": hora, "laptop_id": fila[COL_LAPTOP_ID], "observacion": "NO_ENTREGA_CONTADA"}
                clave = f"{matricula}:{fila[COL_LAPTOP_ID]}:{fila[COL_FECHA]}:{fila[COL_HORA_INGRESO]}"
                crear_y_enviar_evento_critico("NO_ENTREGA", matricula, fila[COL_NOMBRE], fecha, hora, fila[COL_LAPTOP_ID], datos, clave)
                return
        return

        for i in reversed(range(len(registros))):
            fila = registros[i]

            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                observacion = fila[COL_OBSERVACION].strip()

                if observacion == "NO_ENTREGA_CONTADA":
                    break

                objetivo_pendiente = _objetivo_estado_intermedio(
                    observacion, "PROCESANDO_NO_ENTREGA"
                )
                if objetivo_pendiente is not None:
                    if not matricula_exenta_no_entrega(matricula):
                        incrementar_no_entregas(matricula, objetivo_pendiente)
                    _finalizar_observacion(i + 1, "NO_ENTREGA_CONTADA")
                    break

                # NO ENTREGA: existe una sesión pendiente todavía no contabilizada.
                if (
                    hora_salida == ""
                    and observacion_es_normal(observacion)
                    and not matricula_exenta_no_entrega(matricula)
                ):
                    numero_fila = i + 1
                    objetivo = _contador_no_entregas_actual(matricula) + 1
                    estado_actual = hoja_registros.get(
                        f"E{numero_fila}:G{numero_fila}"
                    )
                    if not estado_actual or len(estado_actual[0]) <= 2:
                        raise ErrorConsultaGoogle("No se pudo releer la no entrega")
                    salida_actual = estado_actual[0][0].strip()
                    observacion_actual = estado_actual[0][2].strip()
                    if observacion_actual == "NO_ENTREGA_CONTADA":
                        break
                    objetivo_existente = _objetivo_estado_intermedio(
                        observacion_actual, "PROCESANDO_NO_ENTREGA"
                    )
                    if objetivo_existente is not None:
                        incrementar_no_entregas(matricula, objetivo_existente)
                        _finalizar_observacion(numero_fila, "NO_ENTREGA_CONTADA")
                        break
                    if salida_actual or not observacion_es_normal(observacion_actual):
                        break

                    estado_intermedio = f"PROCESANDO_NO_ENTREGA:{objetivo}"
                    raise RuntimeError("NO_ENTREGA_DEBE_USAR_COORDINADOR")
                    verificacion = hoja_registros.get(f"E{numero_fila}:G{numero_fila}")
                    if (
                        not verificacion
                        or len(verificacion[0]) <= 2
                        or verificacion[0][0].strip() != "-----"
                        or verificacion[0][2].strip() not in (
                            estado_intermedio,
                            "NO_ENTREGA_CONTADA",
                        )
                    ):
                        raise ErrorConsultaGoogle("No se pudo verificar el procesamiento de no entrega")
                    if verificacion[0][2].strip() == "NO_ENTREGA_CONTADA":
                        break
                    incrementar_no_entregas(matricula, objetivo)
                    _finalizar_observacion(numero_fila, "NO_ENTREGA_CONTADA")
                break

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al procesar no entrega: {e}")
        raise ErrorConsultaGoogle("No se pudo procesar la no entrega") from e


def sesion_activa_en_esta_laptop(matricula):
    try:
        if hoja_registros is None:
            return False

        laptop_actual = socket.gethostname()
        registros = hoja_registros.get_all_values()

        for fila in reversed(registros):
            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                return (
                    fila[COL_LAPTOP_ID] == laptop_actual and
                    fila[COL_HORA_SALIDA].strip() == "" and
                    observacion_es_normal(fila[COL_OBSERVACION])
                )

    except Exception as e:
        print(f"Error al validar sesión activa local: {e}")

    return False




def _existe_sesion_activa_para_matricula(registros, matricula):
    return any(
        len(fila) > COL_OBSERVACION
        and str(fila[COL_MATRICULA]).strip() == str(matricula).strip()
        and str(fila[COL_HORA_SALIDA]).strip() == ""
        and observacion_es_normal(fila[COL_OBSERVACION])
        for fila in registros
    )


def registrar_entrada(matricula):
    try:
        registrar_evento_tecnico(f"PRESTAMO_LOCAL_INICIANDO matricula={matricula}")
        if hoja_registros is None:
            registrar_evento_tecnico("PRESTAMO_LOCAL_ERROR tipo=RuntimeError mensaje=HOJA_REGISTROS_NO_DISPONIBLE")
            return None

        nombre = buscar_nombre(matricula)
        if nombre:
            hora, fecha = obtener_hora_internet()
            laptop_id = socket.gethostname()

            with bloqueo_local_registro_prestamo():
                registros = hoja_registros.get_all_values()
                if _existe_sesion_activa_para_matricula(registros, matricula):
                    registrar_evento_tecnico(
                        "PRESTAMO_LOCAL_SESION_EXISTENTE origen=REGISTROS"
                    )
                    return None
                previa = almacen_contingencia.obtener_sesion_activa(laptop_id)
                if previa:
                    registrar_evento_tecnico(
                        "PRESTAMO_LOCAL_SESION_EXISTENTE "
                        f"sesion={previa.get('id_sesion', '')} "
                        f"prestamo_sincronizado={previa.get('prestamo_sincronizado', False)}"
                    )
                sesion = almacen_contingencia.registrar_prestamo_coordinador(
                    matricula, nombre, laptop_id, fecha, hora, socket.gethostname()
                )
                id_evento = sesion.get("prestamo_evento_id")
                if sesion.get("sesion_existente"):
                    registrar_evento_tecnico(
                        "PRESTAMO_LOCAL_SESION_EXISTENTE "
                        f"sesion={sesion.get('id_sesion', '')} "
                        f"reconciliada={sesion.get('sesion_reconciliada', False)}"
                    )
                else:
                    registrar_evento_tecnico(
                        "PRESTAMO_LOCAL_SESION_CREADA "
                        f"sesion={sesion.get('id_sesion', '')} "
                        f"reconciliada_previa={sesion.get('sesion_reconciliada', False)}"
                    )
                if id_evento:
                    if not sesion.get("sesion_existente"):
                        registrar_evento_tecnico(
                            f"PRESTAMO_LOCAL_EVENTO_CREADO id={id_evento} tipo=PRESTAMO"
                        )
                    # registrar_entrada ya corre fuera de Tk. Da al coordinador
                    # una oportunidad breve antes de mostrar la entrega; si falla,
                    # SQLite conserva exactamente el mismo UUID para la cola.
                    sincronizar_evento_inmediato(id_evento)
                else:
                    registrar_evento_tecnico(
                        "PRESTAMO_LOCAL_SESION_EXISTENTE evento=NO_APLICA"
                    )

            return nombre
        registrar_evento_tecnico(
            "PRESTAMO_LOCAL_ERROR tipo=LookupError mensaje=NOMBRE_NO_DISPONIBLE"
        )

    except Exception as e:
        registrar_evento_tecnico(
            f"PRESTAMO_LOCAL_ERROR tipo={type(e).__name__} mensaje={str(e)[:300]}"
        )
        print(f"Error al registrar entrada: {e}")

    return None


def _existe_entrada_del_mismo_intento(
    registros, matricula, nombre, fecha, hora_ingreso, laptop_id
):
    for fila in reversed(registros):
        if (
            len(fila) > COL_OBSERVACION
            and fila[COL_MATRICULA] == matricula
            and fila[COL_NOMBRE] == nombre
            and fila[COL_FECHA] == fecha
            and fila[COL_HORA_INGRESO] == hora_ingreso
            and fila[COL_LAPTOP_ID] == laptop_id
            and fila[COL_HORA_SALIDA].strip() == ""
            and observacion_es_normal(fila[COL_OBSERVACION])
        ):
            return True
    return False


def registrar_salida_con_reintentos(
    nombre, matricula, max_reintentos=3, callback_intento=None,
    solicitar_sincronizacion=True,
):
    # No se confirmó Hora_Salida: el evento durable queda pendiente y no se apaga.
    del nombre, max_reintentos
    if callback_intento:
        callback_intento(1, 1)
    try:
        fecha, hora = _fecha_hora_local()
        sesion_previa = obtener_sesion_local_activa(matricula)
        if not sesion_previa:
            registrar_evento_tecnico(
                "SESION_LOCAL_CIERRE_IGNORADO id_sesion=NO_ENCONTRADA "
                "motivo=DEVOLUCION_SIN_SESION_ACTIVA"
            )
            return False
        registrar_evento_tecnico(
            "SESION_LOCAL_CIERRE_SOLICITADO "
            f"id_sesion={sesion_previa['id_sesion']} tipo_evento=DEVOLUCION"
        )
        sesion = almacen_contingencia.registrar_devolucion_offline(
            matricula, socket.gethostname(), fecha, hora,
            id_sesion=sesion_previa["id_sesion"],
        )
        registrar_evento_tecnico(
            f"SESION_LOCAL_CERRADA id_sesion={sesion['id_sesion']}"
        )
        id_evento = sesion.get("devolucion_evento_id")
        # La UI de entrega puede reservar este único evento para el intento
        # breve pre-apagado; los demás llamadores conservan el daemon normal.
        if id_evento and solicitar_sincronizacion:
            enviar_evento_local_inmediato(id_evento)
        return sesion
    except Exception as error:
        registrar_evento_tecnico(f"COORDINADOR_DEVOLUCION_ERROR={type(error).__name__}")
        return False


def crear_control_devolucion(ventana, matricula, nombre, boton):
    """Construye la única entrada de la devolución para una ventana de entrega.

    ``estado`` es la única autoridad de etapa. Las referencias al modal y a los
    temporizadores no representan etapas ni se usan como guards de negocio.
    """
    estado = {"estado": "IDLE", "timers": {}, "modal": None}
    ventana._estado_devolucion = estado

    def cambiar_estado(nuevo):
        anterior = estado["estado"]
        if anterior == nuevo:
            return
        estado["estado"] = nuevo
        registrar_evento_tecnico(f"DEVOLUCION_ESTADO {anterior}→{nuevo}")

    def cancelar_timer(nombre_timer):
        after_id = estado["timers"].pop(nombre_timer, None)
        if after_id:
            try:
                ventana.after_cancel(after_id)
            except tk.TclError:
                pass

    def cancelar_timers():
        for nombre_timer in tuple(estado["timers"]):
            cancelar_timer(nombre_timer)

    def cerrar_modal():
        modal = estado.get("modal")
        estado["modal"] = None
        if modal:
            cerrar_modal_verificacion_ubicacion(modal)

    def actualizar_modal(titulo, mensaje, texto_estado):
        modal = estado.get("modal")
        if not _widget_tk_vivo(modal):
            modal_progreso = crear_modal_verificacion_ubicacion(
                ventana, mensaje, titulo=titulo, estado=texto_estado
            )
            estado["modal"] = modal_progreso[0]
            return
        # El modal existente no necesita recrearse para cambiar de etapa.
        try:
            for widget in modal.winfo_children():
                if isinstance(widget, tk.Label):
                    texto = widget.cget("text")
                    if texto in ("Verificando ubicación...", "Procesando...", "Guardando devolución...", "Registrando entrega..."):
                        widget.config(text=texto_estado)
        except tk.TclError:
            pass

    def ejecutar_en_ui_devolucion(callback, *args):
        """Devuelve cada resultado worker a Tk y recupera sus excepciones."""
        def protegido():
            try:
                callback(*args)
            except Exception as error:
                registrar_evento_tecnico(
                    f"DEVOLUCION_CALLBACK_ERROR tipo={type(error).__name__}"
                )
                abortar_devolucion("ERROR_CALLBACK_ANTES_DE_SQLITE")

        if threading.current_thread() is threading.main_thread():
            protegido()
            return
        try:
            ventana.after(0, protegido)
        except (AttributeError, tk.TclError):
            registrar_evento_tecnico("DEVOLUCION_CALLBACK_ERROR tipo=TK_NO_DISPONIBLE")

    def abortar_devolucion(motivo):
        """Única salida recuperable: deja el botón listo para otro intento."""
        if estado["estado"] in ("IDLE", "FINALIZADA", "SINCRONIZANDO"):
            return
        registrar_evento_tecnico(f"DEVOLUCION_ABORTADA motivo={motivo}")
        cancelar_timers()
        cerrar_modal()
        cambiar_estado("ERROR_RECUPERABLE")
        try:
            boton.config(state="normal", text="ENTREGAR Y APAGAR")
        except tk.TclError:
            pass
        # El command se configura una única vez al crear el botón y nunca se
        # reemplaza durante el intento.
        cambiar_estado("IDLE")

    def finalizar_devolucion():
        if estado["estado"] != "SINCRONIZANDO":
            return
        cancelar_timers()
        cambiar_estado("FINALIZADA")
        registrar_evento_tecnico("DEVOLUCION_FINALIZADA")
        cerrar_modal()
        try:
            ventana.destroy()
        except tk.TclError:
            pass
        apagar_windows()

    def watchdog_pre_durable():
        if estado["estado"] in ("VALIDANDO_PROXIMIDAD", "PERSISTIENDO_LOCAL"):
            abortar_devolucion("TIMEOUT_PRE_DURABLE")

    def watchdog_post_durable():
        if estado["estado"] == "SINCRONIZANDO":
            registrar_evento_tecnico("DEVOLUCION_WATCHDOG_APAGADO")
            finalizar_devolucion()

    def iniciar_sincronizacion(id_evento):
        if estado["estado"] != "PERSISTIENDO_LOCAL":
            return
        cancelar_timer("pre_durable")
        cambiar_estado("SINCRONIZANDO")
        registrar_evento_tecnico(f"DEVOLUCION_SYNC_INMEDIATA_INICIADA id={id_evento}")
        estado["timers"]["post_durable"] = ventana.after(6000, watchdog_post_durable)

        def worker_sincronizacion():
            try:
                aplicada = sincronizar_devolucion_pre_apagado(id_evento)
                registrar_evento_tecnico(
                    "DEVOLUCION_SYNC_APLICADA" if aplicada else "DEVOLUCION_SYNC_PENDIENTE"
                )
                ejecutar_en_ui_devolucion(finalizar_devolucion)
            except Exception as error:
                registrar_evento_tecnico(
                    f"DEVOLUCION_SYNC_ERROR tipo={type(error).__name__}"
                )
                # SQLite ya confirmó: el evento queda pendiente y se apaga.
                ejecutar_en_ui_devolucion(finalizar_devolucion)

        threading.Thread(target=worker_sincronizacion, daemon=True).start()

    def procesar_persistencia(sesion):
        if estado["estado"] != "PERSISTIENDO_LOCAL":
            return
        if not sesion:
            abortar_devolucion("SQLITE_NO_CONFIRMADA")
            return
        id_evento = sesion.get("devolucion_evento_id")
        if not id_evento:
            abortar_devolucion("EVENTO_DEVOLUCION_AUSENTE")
            return
        # finalizar_devolucion_durable sólo retorna tras cerrar el id_sesion
        # exacto y dejar el UUID de DEVOLUCION en SQLite.
        iniciar_sincronizacion(id_evento)

    def persistir_devolucion():
        try:
            sesion = finalizar_devolucion_durable(
                matricula, solicitar_sincronizacion=False
            )
            ejecutar_en_ui_devolucion(procesar_persistencia, sesion)
        except Exception as error:
            registrar_evento_tecnico(
                f"DEVOLUCION_PERSISTENCIA_ERROR tipo={type(error).__name__}"
            )
            ejecutar_en_ui_devolucion(abortar_devolucion, "ERROR_ANTES_DE_SQLITE")

    def iniciar_persistencia(autorizacion_administrativa=False, estado_proximidad=None):
        if estado["estado"] != "VALIDANDO_PROXIMIDAD":
            return
        if autorizacion_administrativa:
            registrar_evento_tecnico(
                "DEVOLUCION_PROXIMIDAD_AUTORIZADA_ADMIN "
                f"estado={estado_proximidad or 'DESCONOCIDO'}"
            )
        cambiar_estado("PERSISTIENDO_LOCAL")
        try:
            boton.config(text="Guardando devolución...")
            actualizar_modal(
                "Registrando devolución",
                "Estamos guardando la entrega del equipo. Por favor, espera.",
                "Guardando devolución...",
            )
            registrar_evento_tecnico("DEVOLUCION_DURABLE_PIPELINE_INICIADO")
            threading.Thread(target=persistir_devolucion, daemon=True).start()
        except Exception as error:
            registrar_evento_tecnico(
                f"DEVOLUCION_TRANSICION_ERROR tipo={type(error).__name__}"
            )
            abortar_devolucion("ERROR_ANTES_DE_SQLITE")

    def procesar_resultado_proximidad(
        proximidad_valida, estado_proximidad, autorizacion_administrativa=False
    ):
        if estado["estado"] != "VALIDANDO_PROXIMIDAD":
            return
        try:
            if proximidad_valida:
                iniciar_persistencia(autorizacion_administrativa, estado_proximidad)
                return

            permite_contingencia = estado_proximidad in (ESTADO_NO_DETECTADO, ESTADO_ERROR)
            contexto = iniciar_contexto_contingencia(
                estado_proximidad,
                "DEVOLUCION",
                matricula,
                ventana,
                # El bypass entra explícitamente a esta misma continuación; no
                # depende de que wait_window devuelva el control.
                reanudar=lambda: procesar_resultado_proximidad(
                    True, estado_proximidad, autorizacion_administrativa=True
                ),
                recuperar_bypass=abortar_devolucion,
            )
            mostrar_fallo_proximidad(
                estado_proximidad,
                parent=ventana,
                al_crear=registrar_dialogo_fallo_contingencia,
            )
            autorizada = finalizar_contexto_contingencia()
            if autorizada:
                if contexto.get("reanudado"):
                    return
                procesar_resultado_proximidad(
                    True, estado_proximidad, autorizacion_administrativa=True
                )
                return
            if not permite_contingencia:
                abortar_devolucion("PROXIMIDAD_NO_AUTORIZADA")
                return
            abortar_devolucion("PROXIMIDAD_NO_AUTORIZADA")
        except Exception as error:
            registrar_evento_tecnico(
                f"DEVOLUCION_PROXIMIDAD_ERROR tipo={type(error).__name__}"
            )
            abortar_devolucion("ERROR_ANTES_DE_SQLITE")

    def worker_proximidad():
        try:
            resultado = verificar_proximidad_carrito(callback_evento=registrar_evento_tecnico)
            ejecutar_en_ui_devolucion(procesar_resultado_proximidad, *resultado)
        except Exception as error:
            registrar_evento_tecnico(
                f"DEVOLUCION_PROXIMIDAD_WORKER_ERROR tipo={type(error).__name__}"
            )
            ejecutar_en_ui_devolucion(abortar_devolucion, "ERROR_ANTES_DE_SQLITE")

    def iniciar_devolucion():
        """Única entrada pública del botón ENTREGAR Y APAGAR."""
        if estado["estado"] != "IDLE":
            registrar_evento_tecnico(
                f"DEVOLUCION_CLIC_IGNORADO estado={estado['estado']}"
            )
            return
        cambiar_estado("VALIDANDO_PROXIMIDAD")
        try:
            boton.config(
                state="disabled",
                text="Verificando ubicación..." if PROXIMIDAD_HABILITADA else "Guardando devolución...",
            )
            estado["modal"] = crear_modal_verificacion_ubicacion(
                ventana,
                "Asegúrate de estar cerca del carrito para continuar con la devolución.",
            )[0]
            estado["timers"]["pre_durable"] = ventana.after(5000, watchdog_pre_durable)
            if PROXIMIDAD_HABILITADA:
                threading.Thread(target=worker_proximidad, daemon=True).start()
            else:
                ventana.after_idle(
                    lambda: procesar_resultado_proximidad(True, None)
                )
        except Exception as error:
            registrar_evento_tecnico(
                f"DEVOLUCION_INICIO_ERROR tipo={type(error).__name__}"
            )
            abortar_devolucion("ERROR_ANTES_DE_SQLITE")

    ventana._abortar_devolucion = abortar_devolucion
    return iniciar_devolucion


def mostrar_ventana_entrega(nombre, matricula):
    global ventana_entrega_activa

    ventana_entrega = tk.Toplevel()
    vincular_dispatcher_administrativo(ventana_entrega)
    ventana_entrega_activa = ventana_entrega

    def limpiar_referencia_entrega(event):
        global ventana_entrega_activa
        if event.widget == ventana_entrega:
            ventana_entrega_activa = None

    ventana_entrega.bind("<Destroy>", limpiar_referencia_entrega)

    
    # CLAVE: ventana independiente
    ventana_entrega.transient(None)
    ventana_entrega.title("Entrega de Laptop")
    ventana_entrega.resizable(False, False)
    ventana_entrega.configure(bg=COLOR_FONDO)

    # ? Bloquear cerrar, pero NO minimizar
    ventana_entrega.protocol("WM_DELETE_WINDOW", lambda: None)

    # Permitir minimizar
    ventana_entrega.minsize(420, 480)

    frame_principal = tk.Frame(
        ventana_entrega,
        bg=COLOR_TARJETA,
        padx=35,
        pady=30
    )
    frame_principal.pack(fill=tk.BOTH, expand=True)

    if TEMA_INTERFAZ == "moderno":
        tk.Frame(frame_principal, height=4, bg=COLOR_EXITO).pack(
            fill=tk.X, pady=(0, 18)
        )

    # Logo
    logo_entrega = cargar_logo(RUTA_UTP, 120, 70)
    if logo_entrega:
        lbl_logo = tk.Label(frame_principal, image=logo_entrega, bg=COLOR_TARJETA)
        lbl_logo.image = logo_entrega
        lbl_logo.pack(pady=(0, 20))


    tk.Label(
        frame_principal,
        text="Sistema de Control de Laptops",
        font=FUENTE_SUBTITULO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 5))

    tk.Label(
        frame_principal,
        text=f"Bienvenid@,\n{nombre}",
        font=("Segoe UI", 16, "bold"),
        fg=COLOR_EXITO,
        bg=COLOR_TARJETA,
        justify="center",
        wraplength=340
    ).pack(pady=(0, 25))

    # Botón principal: su command se asigna una sola vez y no cambia durante
    # ningún intento de devolución.
    btn_entregar = tk.Button(
        frame_principal,
        text="ENTREGAR Y APAGAR",
        padx=30,
        pady=15,
        **estilo_boton("error")
    )
    iniciar_devolucion = crear_control_devolucion(
        ventana_entrega, matricula, nombre, btn_entregar
    )
    btn_entregar.config(command=iniciar_devolucion)
    btn_entregar.pack(pady=20)

    tk.Label(
        frame_principal,
        text="La laptop se apagará automáticamente",
        font=FUENTE_PEQ,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(pady=(10, 0))



    # Centrar ventana
    ancho, alto = 420, 480
    x = (ventana_entrega.winfo_screenwidth() - ancho) // 2
    y = (ventana_entrega.winfo_screenheight() - alto) // 2
    ventana_entrega.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana_entrega.deiconify()
    ventana_entrega.lift()
    ventana_entrega.focus_force()

def formatear_fecha(fecha_entrada):
    try:
        fecha_obj = datetime.strptime(fecha_entrada, "%Y-%m-%d")
        meses_espanol = {
            1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
            5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
            9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
        }
        dia = fecha_obj.day
        mes = meses_espanol[fecha_obj.month]
        return f"{dia} de {mes}"
    except:
        return fecha_entrada

def mostrar_aviso_entrega_pendiente(fecha_entrada, laptop_id):
    fecha_formateada = formatear_fecha(fecha_entrada)
    
    mensaje = f"Olvidaste entregar la laptop\n\n"
    mensaje += f"Laptop: {laptop_id}\n"
    mensaje += f"Fecha: {fecha_formateada}\n\n"
    mensaje += "Recuerda siempre usar el botón 'Entregar y Apagar'"
    
    mostrar_advertencia("Entrega Pendiente", mensaje)
def verificar_sesion_activa_en_otra_laptop(matricula):
    """
    Verifica si la matrícula tiene una sesión activa
    en OTRA laptop diferente a la actual.
    """
    try:
        if hoja_registros is None:
            raise ErrorConsultaGoogle("La hoja Registros no está disponible")

        laptop_actual = socket.gethostname()
        registros = hoja_registros.get_all_values()

        for fila in reversed(registros):
            if len(fila) > COL_OBSERVACION and fila[0] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                laptop_registro = fila[COL_LAPTOP_ID]
                observacion = fila[COL_OBSERVACION].strip()


                if (
                    hora_salida == ""
                    and observacion_es_normal(observacion)
                    and laptop_registro != laptop_actual
                ):
                    return True, laptop_registro

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al verificar sesión en otra laptop: {e}")
        raise ErrorConsultaGoogle("No se pudo verificar la sesión activa") from e

    return False, None


def mostrar_confirmacion_simple(nombre, matricula, identidad_confirmada=False,
                                modo_contingencia=False, persona_cache=None):
    global entrada, ventana, procesando_sesion

    if not identidad_confirmada:
        activar_timeout_inactividad()
        respuesta = mostrar_confirmacion_personalizada(
            "Confirmación",
            f"¿Eres {nombre}?"
        )
        cancelar_timeout_inactividad()

        if not respuesta:
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return False

    cancelar_timeout_inactividad()
    cambiar_estado("Consultando sesión...", COLOR_ADVERTENCIA)

    def consultar_sesion():
        try:
            resultado = verificar_sesion_activa_en_otra_laptop(matricula)
            ejecutar_en_ui(procesar_sesion, "ok", *resultado)
        except ErrorConsultaGoogle:
            ejecutar_en_ui(procesar_sesion, "error", False, None)

    def resolver_sesion_local():
        previa = obtener_sesion_local_activa()
        if not previa:
            ejecutar_en_ui(procesar_sesion_local, "sin_sesion", None)
            return
        registrar_evento_tecnico(
            f"SESION_RESIDUAL_DETECTADA id_sesion={previa['id_sesion']}"
        )
        try:
            fecha, hora = _fecha_hora_local()
            resultado = almacen_contingencia.cerrar_sesion_residual_al_arrancar(
                socket.gethostname(), fecha, hora, socket.gethostname()
            )
            cerrada = resultado.get("sesion")
            if resultado.get("resultado") == "SESION_RESIDUAL_CERRADA" and cerrada:
                if resultado.get("no_entrega_creada"):
                    registrar_evento_tecnico(
                        f"NO_ENTREGA_RESIDUAL_CREADA id_sesion={cerrada['id_sesion']}"
                    )
                for campo in ("no_entrega_evento_id", "cierre_automatico_evento_id"):
                    if cerrada.get(campo):
                        enviar_evento_local_inmediato(cerrada[campo])
                registrar_evento_tecnico(
                    f"SESION_RESIDUAL_CERRADA id_sesion={cerrada['id_sesion']}"
                )
                ejecutar_en_ui(procesar_sesion_local, "residual_cerrada", cerrada)
                return
        except Exception as error:
            registrar_evento_tecnico(
                f"SESION_RESIDUAL_CIERRE_ERROR tipo={type(error).__name__}"
            )
        ejecutar_en_ui(procesar_sesion_local, "error", previa)

    def procesar_sesion_local(resultado, sesion):
        global procesando_sesion
        if resultado == "error":
            mostrar_error(
                "Error",
                "No se pudo resolver la sesión anterior.\n\nIntenta nuevamente."
            )
            reiniciar_estado_sistema()
            return
        registrar_evento_tecnico("FLUJO_SELECCIONADO tipo=PRESTAMO_NUEVO")
        iniciar_validacion_identidad()

    def procesar_sesion(resultado_consulta, sesion_activa, laptop_otro):
        if resultado_consulta == "error":
            mostrar_error(
                "No se pudo verificar la información",
                "No fue posible verificar la conexión. Intenta nuevamente."
            )
            reiniciar_estado_sistema()
            return

        if sesion_activa:
            activar_timeout_inactividad()
            continuar = mostrar_confirmacion_personalizada(
                "Sesión activa detectada",
                "Se detectó una sesión activa en otro equipo.\n\n"
                f"Laptop anterior: {laptop_otro}\n\n"
                "Si continúas:\n"
                "• La sesión anterior se cerrará automáticamente\n"
                "• Se registrará como NO ENTREGA\n"
                "• Esta acción quedará registrada\n\n"
                "¿Deseas continuar en este equipo?"
            )
            cancelar_timeout_inactividad()
            if not continuar:
                reiniciar_estado_sistema()
                entrada.delete(0, tk.END)
                entrada.focus()
                return

            cambiar_estado("Actualizando sesión anterior...", COLOR_ADVERTENCIA)
            threading.Thread(target=cerrar_sesion_anterior, daemon=True).start()
            return

        threading.Thread(target=resolver_sesion_local, daemon=True).start()

    def cerrar_sesion_anterior():
        try:
            exito = cerrar_sesion_anterior_y_contar_no_entrega(matricula)
        except ErrorConsultaGoogle:
            exito = False
        ejecutar_en_ui(procesar_cierre_anterior, exito)

    def procesar_cierre_anterior(exito):
        if not exito:
            mostrar_error(
                "Error",
                "No se pudo cerrar la sesión anterior.\n\nIntenta nuevamente."
            )
            reiniciar_estado_sistema()
            return
        threading.Thread(target=resolver_sesion_local, daemon=True).start()

    def iniciar_validacion_identidad():
        cambiar_estado("Validando identidad...", COLOR_ADVERTENCIA)
        threading.Thread(target=consultar_identidad, daemon=True).start()

    def consultar_identidad():
        try:
            if modo_contingencia:
                rol = normalizar_rol(persona_cache.get("rol"))
                curp_real = (
                    persona_cache.get("curp_ultimos_2")
                    if rol_requiere_curp(rol) else None
                )
            else:
                rol, curp_real = consultar_datos_identidad(matricula)
            if rol is None:
                registrar_evento_tecnico(
                    f"IDENTIDAD matricula={matricula} resultado=NO_ENCONTRADA"
                )
                ejecutar_en_ui(procesar_identidad, "no_encontrada", None, None)
                return
            flujo = seleccionar_flujo_identidad(rol)
            registrar_evento_tecnico(
                f"IDENTIDAD matricula={matricula} encontrada=SI "
                f"rol={rol} flujo={flujo} "
                f"curp={'CORRESPONDE' if rol_requiere_curp(rol) else 'OMITIDA_POR_ROL'}"
            )
            ejecutar_en_ui(procesar_identidad, "ok", rol, curp_real)
        except ErrorConsultaGoogle as error:
            causa = error.__cause__ or error
            registrar_evento_tecnico(
                f"IDENTIDAD matricula={matricula} error_consulta="
                f"{type(causa).__name__}: {causa}"
            )
            ejecutar_en_ui(procesar_identidad, "error_consulta", None, None)

    def procesar_identidad(resultado, rol, curp_real):
        def cancelar_validacion_curp():
            cancelar_timeout_inactividad()
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus_set()

        if resultado == "error_consulta":
            mostrar_error(
                "No se pudo verificar la información",
                "No fue posible verificar la conexión. Intenta nuevamente."
            )
            cancelar_validacion_curp()
            return

        if resultado == "no_encontrada":
            mostrar_error(
                "Matrícula no válida",
                "La matrícula no está registrada.\n\nVerifica o contacta al administrador."
            )
            cancelar_validacion_curp()
            return

        if not rol_requiere_curp(rol):
            registrar_evento_tecnico(
                f"IDENTIDAD matricula={matricula} rol={rol} CURP_OMITIDA_POR_ROL"
            )
            continuar_despues_identidad()
            return

        if not curp_real:
            registrar_evento_tecnico(
                f"IDENTIDAD matricula={matricula} rol={rol} CURP_ALUMNO_VACIA"
            )
            if modo_contingencia:
                mostrar_error(
                    "Validación offline no disponible",
                    "Esta matrícula está en la caché local, pero todavía no tiene "
                    "el verificador de identidad guardado.\n\n"
                    "Conéctate a Internet una vez para actualizarla."
                )
                cancelar_validacion_curp()
                return
            mostrar_error(
                "Error",
                "No se pudo validar la identidad.\nContacta al administrador."
            )
            cancelar_validacion_curp()
            return

        for intento in range(1, 4):
            registrar_evento_tecnico(
                f"CURP_SOLICITADA matricula={matricula} intento={intento}"
            )
            curp_ingresada = pedir_curp_ultimos_2(ventana)
            if not curp_ingresada:
                cancelar_validacion_curp()
                return

            if validar_curp_ultimos_2(curp_real, curp_ingresada):
                break

            intentos_restantes = 3 - intento
            if intentos_restantes:
                texto_intentos = (
                    "Te queda 1 intento."
                    if intentos_restantes == 1
                    else f"Te quedan {intentos_restantes} intentos."
                )
                mostrar_error(
                    "Datos incorrectos",
                    "Los datos ingresados no coinciden.",
                    texto_destacado=texto_intentos,
                    texto_boton="INTENTAR NUEVAMENTE"
                )
                if flujo_cancelado_por_inactividad:
                    cancelar_validacion_curp()
                    return
                continue

            mostrar_error(
                "Acceso denegado",
                "Los datos no coinciden.\n\n"
                "Se agotaron los intentos de validación. Sesión cancelada."
            )
            cancelar_validacion_curp()
            return

        registrar_evento_tecnico(
            f"IDENTIDAD matricula={matricula} rol={rol} CURP_VALIDADA"
        )
        continuar_despues_identidad()

    def continuar_despues_identidad():
        cancelar_timeout_inactividad()
        cambiar_estado("Consultando registros...", COLOR_ADVERTENCIA)
        threading.Thread(target=consultar_control, daemon=True).start()

    def consultar_control():
        try:
            if modo_contingencia:
                no_entregas = no_entregas_efectivas_local(
                    matricula, persona_cache.get("no_entregas")
                )
                estado = persona_cache.get("estado")
                tiene_pendiente = False
                fecha_entrada = None
                laptop_id = None
            else:
                procesar_no_entrega_si_corresponde(matricula)
                no_entregas, estado = obtener_control_alumno(matricula)
                tiene_pendiente, fecha_entrada, laptop_id = verificar_entrega_pendiente(matricula)
            ejecutar_en_ui(
                procesar_control,
                no_entregas,
                estado,
                tiene_pendiente,
                fecha_entrada,
                laptop_id
            )
        except ErrorConsultaGoogle:
            ejecutar_en_ui(procesar_error_control)

    def procesar_error_control():
        mostrar_error(
            "No se pudo verificar la información",
            "No fue posible verificar la conexión. Intenta nuevamente."
        )
        reiniciar_estado_sistema()

    def procesar_control(no_entregas, estado, tiene_pendiente, fecha_entrada, laptop_id):
        estado = normalizar_estado_alumno(estado)
        if estado is None:
            registrar_evento_tecnico(
                "ESTADO_ALUMNO_BLOQUEADO valor=DESCONOCIDO"
            )
            mostrar_error(
                "Acceso denegado",
                "El estado del alumno no permite registrar un préstamo.\n\n"
                "Contacta al administrador."
            )
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return

        if no_entregas >= 2:
            activar_timeout_inactividad()
            mostrar_ventana_control_unificada(
                matricula=matricula,
                nombre=nombre,
                tiene_pendiente=tiene_pendiente,
                fecha_entrada=fecha_entrada,
                laptop_id=laptop_id,
                no_entregas=no_entregas,
                estado=("" if estado == ESTADO_VACIO_VALIDO else estado)
            )
            cancelar_timeout_inactividad()
            if flujo_cancelado_por_inactividad:
                reiniciar_estado_sistema()
                return
            if estado == "SANCIONADO":
                reiniciar_estado_sistema()
                entrada.delete(0, tk.END)
                entrada.focus()
                return

        if not estado_alumno_permite_prestamo(estado):
            registrar_evento_tecnico(
                f"ESTADO_ALUMNO_BLOQUEADO valor={estado}"
            )
            mostrar_error(
                "Acceso denegado",
                "El estado del alumno no permite registrar un préstamo.\n\n"
                "Contacta al administrador."
            )
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return

        cambiar_estado("Registrando préstamo...", COLOR_ADVERTENCIA)
        threading.Thread(target=registrar_prestamo, daemon=True).start()

    def registrar_prestamo():
        if modo_contingencia:
            resultado = nombre if registrar_prestamo_contingencia(matricula, nombre) else None
        else:
            resultado = registrar_entrada(matricula)
        ejecutar_en_ui(finalizar_prestamo, resultado)

    def finalizar_prestamo(resultado):
        global procesando_sesion
        if not resultado:
            mostrar_error("Error", "No se pudo registrar la entrada.")
            reiniciar_estado_sistema()
            return

        procesando_sesion = False
        ventana.withdraw()
        ventana.after(200, lambda: mostrar_ventana_entrega(nombre, matricula))

    if modo_contingencia:
        threading.Thread(target=resolver_sesion_local, daemon=True).start()
    else:
        threading.Thread(target=consultar_sesion, daemon=True).start()
    return True


def cambiar_estado(texto, color=None):
    if threading.current_thread() is not threading.main_thread():
        ejecutar_en_ui(cambiar_estado, texto, color)
        return

    estado_var.set(texto)
    if color:
        estado_label.config(fg=color)
    ventana.update_idletasks()

def reiniciar_estado_sistema():
    """Reinicia el estado del sistema después de una operación"""
    global procesando_sesion
    procesando_sesion = False
    btn_entrar.config(state="normal")
    # Actualizar estado de conexión actual
    if verificar_conexion_base_datos():
        cambiar_estado("Conectado", COLOR_EXITO)
    else:
        cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        if acepta_estado_equipo:
            acepta_estado_equipo.set(False)
            try:
                pass
            except:
                pass



def mostrar_aviso_internet_bloqueante():
    """Muestra un aviso emergente bloqueante que no se puede mover, cerrar ni minimizar"""
    global aviso_internet

    try:
        if aviso_internet and aviso_internet.winfo_exists():
            aviso_internet.lift()
            return
    except:
        aviso_internet = None
    aviso_internet = tk.Toplevel(ventana)
    vincular_dispatcher_administrativo(aviso_internet)
    aviso_internet.title("Conexión Requerida")
    
    # Hacer la ventana completamente bloqueante
    aviso_internet.attributes('-topmost', True)
    aviso_internet.resizable(False, False)
    aviso_internet.overrideredirect(True)
    
    # Configurar tamaño y posición
    ancho_aviso = 500
    alto_aviso = 320
    
    pantalla_ancho = aviso_internet.winfo_screenwidth()
    pantalla_alto = aviso_internet.winfo_screenheight()
    x = (pantalla_ancho - ancho_aviso) // 2
    y = (pantalla_alto - alto_aviso) // 2
    
    aviso_internet.geometry(f"{ancho_aviso}x{alto_aviso}+{x}+{y}")
    aviso_internet.configure(bg=COLOR_PRIMARIO)
    aviso_internet.focus_force()
    
    # Frame principal
    frame_aviso = tk.Frame(aviso_internet, bg=COLOR_TARJETA, bd=0, relief=tk.FLAT)
    frame_aviso.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
    
    # Logo
    logo_aviso = cargar_logo(RUTA_UTP, 70, 45)
    if logo_aviso:
        lbl_logo_aviso = tk.Label(frame_aviso, image=logo_aviso, bg=COLOR_TARJETA)
        lbl_logo_aviso.image = logo_aviso
        lbl_logo_aviso.pack(pady=(25, 10))
    
    # Icono de advertencia
    tk.Label(frame_aviso, 
             text="!",
             font=("Segoe UI", 32),
             bg=COLOR_TARJETA).pack(pady=(10, 15))
    
    # Mensaje principal
    tk.Label(frame_aviso, 
             text="NO SE HA DETECTADO CONEXIÓN A INTERNET",
             font=("Segoe UI", 13, "bold"),
             fg=COLOR_ERROR,
             bg=COLOR_TARJETA).pack(pady=(0, 10))
    
    # Instrucciones
    tk.Label(frame_aviso, 
             text="Para utilizar el sistema es necesario contar con conexión a internet.",
             font=FUENTE_CUERPO,
             fg=COLOR_TEXTO,
             bg=COLOR_TARJETA,
             wraplength=400,
             justify=tk.CENTER).pack(pady=(0, 5))
    
    tk.Label(frame_aviso, 
             text="Conecte la computadora a internet para continuar.",
             font=FUENTE_CUERPO_BOLD,
             fg=COLOR_TEXTO,
             bg=COLOR_TARJETA,
             wraplength=400,
             justify=tk.CENTER).pack(pady=(0, 20))
    
    # Barra de progreso animada
    progress_frame = tk.Frame(frame_aviso, bg=COLOR_TARJETA)
    progress_frame.pack(pady=15)
    
    progress_bar = tk.Frame(progress_frame, height=8, bg=COLOR_BORDE, width=350)
    progress_bar.pack()
    progress_bar_inner = tk.Frame(progress_bar, height=8, bg=COLOR_PRIMARIO, width=0)
    progress_bar_inner.place(relx=0, rely=0, relheight=1)
    
    # Mensaje de estado
    estado_label = tk.Label(frame_aviso, 
                           text="Verificando conexión...",
                           font=FUENTE_PEQ,
                           fg=COLOR_TEXTO_SECUNDARIO,
                           bg=COLOR_TARJETA)
    estado_label.pack(pady=5)
    
    # Contador de tiempo
    tiempo_inicio = time.time()
    tiempo_label = tk.Label(frame_aviso, 
                           text="Tiempo esperando: 0 segundos",
                           font=FUENTE_PEQ,
                           fg=COLOR_TEXTO_SECUNDARIO,
                           bg=COLOR_TARJETA)
    tiempo_label.pack(pady=2)
    
    # Función para animar la barra de progreso
    def animar_barra_progreso():
        nonlocal progreso_actual, direccion
        if direccion == 1:
            progreso_actual += 5
            if progreso_actual >= 350:
                progreso_actual = 350
                direccion = -1
        else:
            progreso_actual -= 5
            if progreso_actual <= 0:
                progreso_actual = 0
                direccion = 1
        
        try:
            if not aviso_internet or not aviso_internet.winfo_exists():
                return
            progress_bar_inner.config(width=progreso_actual)
            aviso_internet.after(50, animar_barra_progreso)
        except:
            return
    
    # Función para verificar internet continuamente sin congelar la ventana
    def verificar_internet_continuamente():
        global verificacion_aviso_en_curso

        if verificacion_aviso_en_curso:
            return

        try:
            if not aviso_internet or not aviso_internet.winfo_exists():
                return
        except:
            return

        tiempo_transcurrido = int(time.time() - tiempo_inicio)
        tiempo_label.config(text=f"Tiempo esperando: {tiempo_transcurrido} segundos")
        estado_label.config(text=f"Verificando conexión... ({time.strftime('%H:%M:%S')})")
        verificacion_aviso_en_curso = True

        def revisar_conexion():
            conectado = verificar_internet()
            if conectado:
                conectado = conectar_google_sheets(internet_verificado=True)
            ejecutar_en_ui(procesar_resultado_conexion, conectado)

        threading.Thread(target=revisar_conexion, daemon=True).start()

    def procesar_resultado_conexion(conectado):
        global verificacion_aviso_en_curso
        verificacion_aviso_en_curso = False

        try:
            if not aviso_internet or not aviso_internet.winfo_exists():
                return
        except:
            return

        if conectado:
            cerrar_aviso_internet()
            cambiar_estado("Conectado", COLOR_EXITO)
        else:
            aviso_internet.after(1000, verificar_internet_continuamente)
    
    # Iniciar animación y verificación
    progreso_actual = 0
    direccion = 1
    animar_barra_progreso()
    aviso_internet.after(500, verificar_internet_continuamente)
def iniciar_sesion():
    global procesando_sesion, flujo_cancelado_por_inactividad
    cancelar_timeout_inactividad()

    if procesando_sesion:
        return

    flujo_cancelado_por_inactividad = False
    procesando_sesion = True
    btn_entrar.config(state="disabled")

    matricula = entrada.get().strip()

    if not matricula:
        mostrar_advertencia("Campo vacío", "Por favor ingresa tu matrícula.")
        entrada.focus()
        reiniciar_estado_sistema()
        return
        # Validar aceptación de estado del equipo
    if not acepta_estado_equipo.get():
        mostrar_advertencia(
            "Confirmación requerida",
            "Debes confirmar el estado de la laptop para continuar."
        )
        reiniciar_estado_sistema()
        return

    cancelar_timeout_inactividad()
    ventana_ubicacion = None
    titulo_ubicacion = None
    mensaje_ubicacion = None
    estado_ubicacion = None
    proximidad_visual_activa = {"valor": False}
    alumno_confirmado = {"nombre": None}
    acceso_contingencia = {"activo": False, "persona": None}


    def consultar_acceso():
        def notificar_intento(intento, total):
            ejecutar_en_ui(
                cambiar_estado,
                "Verificando conexión...",
                COLOR_ADVERTENCIA
            )

        if not verificar_sheets_con_reintentos(
            max_intentos=3, callback_intento=notificar_intento
        ):
            persona = obtener_persona_cache(matricula)
            if persona:
                ejecutar_en_ui(procesar_acceso_seguro, "contingencia", persona)
            else:
                ejecutar_en_ui(procesar_acceso_seguro, "sin_cache", None)
            return

        nombre = buscar_nombre(matricula)
        ejecutar_en_ui(procesar_acceso_seguro, "ok", nombre)

    def procesar_acceso_seguro(estado, nombre):
        try:
            procesar_acceso(estado, nombre)
        except Exception:
            traceback.print_exc()
            reiniciar_estado_sistema()

    def procesar_acceso(estado, nombre):
        nonlocal ventana_ubicacion, titulo_ubicacion
        nonlocal mensaje_ubicacion, estado_ubicacion
        if estado == "sin_cache":
            mostrar_error(
                "Validación no disponible",
                "No es posible validar esta matrícula en este momento.\n\n"
                "Intenta nuevamente más tarde."
            )
            reiniciar_estado_sistema()
            return

        if estado == "contingencia":
            acceso_contingencia["activo"] = True
            acceso_contingencia["persona"] = nombre
            nombre = nombre.get("nombre")
            cambiar_estado("Modo de contingencia", COLOR_ADVERTENCIA)
            mostrar_informacion(
                "Modo de contingencia",
                "El registro se guardará y sincronizará automáticamente."
            )

        if not nombre:
            mostrar_error(
                "Matrícula no válida",
                "La matrícula no está registrada.\n\nVerifica o contacta al administrador."
            )
            entrada.delete(0, tk.END)
            entrada.focus()
            reiniciar_estado_sistema()
            return

        registrar_evento_tecnico(
            f"IDENTIDAD matricula={matricula} encontrada=SI etapa=NOMBRE"
        )
        activar_timeout_inactividad()
        confirmado = mostrar_confirmacion_personalizada(
            "Confirmación",
            f"¿Eres {nombre}?"
        )
        cancelar_timeout_inactividad()
        if not confirmado:
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return

        alumno_confirmado["nombre"] = nombre
        if not PROXIMIDAD_HABILITADA:
            mostrar_confirmacion_simple(
                alumno_confirmado["nombre"],
                matricula,
                identidad_confirmada=True,
                modo_contingencia=acceso_contingencia["activo"],
                persona_cache=acceso_contingencia["persona"]
            )
            return

        (
            ventana_ubicacion,
            titulo_ubicacion,
            mensaje_ubicacion,
            estado_ubicacion
        ) = crear_modal_verificacion_ubicacion(
            ventana,
            "Asegúrate de estar cerca del carrito para continuar."
        )
        proximidad_visual_activa["valor"] = True
        threading.Thread(target=comprobar_proximidad, daemon=True).start()

    def comprobar_proximidad():
        try:
            proximidad_valida, estado_proximidad = verificar_proximidad_carrito(
                callback_intento=notificar_intento_ubicacion,
                callback_evento=registrar_evento_tecnico
            )
        except Exception:
            proximidad_valida = False
            estado_proximidad = ESTADO_ERROR

        ejecutar_en_ui(
            procesar_proximidad,
            proximidad_valida,
            estado_proximidad
        )

    def actualizar_intento_ubicacion(intento, total):
        if proximidad_visual_activa["valor"]:
            try:
                estado_ubicacion.config(text=f"Intento {intento} de {total}")
            except tk.TclError:
                pass

    def notificar_intento_ubicacion(intento, total):
        ejecutar_en_ui(actualizar_intento_ubicacion, intento, total)

    def procesar_proximidad(
        proximidad_valida, estado_proximidad, autorizacion_administrativa=False
    ):
        proximidad_visual_activa["valor"] = False

        if not proximidad_valida:
            permite_contingencia = estado_proximidad in (
                ESTADO_NO_DETECTADO,
                ESTADO_ERROR
            )
            contexto = iniciar_contexto_contingencia(
                estado_proximidad,
                "PRESTAMO_NUEVO",
                matricula,
                ventana,
                reanudar=lambda: procesar_proximidad(
                    True, estado_proximidad, autorizacion_administrativa=True
                )
            )
            if not permite_contingencia:
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)

            try:
                activar_timeout_inactividad()
                mostrar_fallo_proximidad(
                    estado_proximidad,
                    parent=ventana,
                    al_crear=registrar_dialogo_fallo_contingencia
                )
            except Exception:
                traceback.print_exc()
            finally:
                cancelar_timeout_inactividad()

            contingencia_autorizada = finalizar_contexto_contingencia()
            if contingencia_autorizada:
                if contexto.get("reanudado"):
                    return
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                procesar_proximidad(True, None, autorizacion_administrativa=True)
                return

            if permite_contingencia:
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
            reiniciar_estado_sistema()
            return

        cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
        registrar_evento_tecnico("FLUJO_PRESTAMO_NUEVO")
        mostrar_confirmacion_simple(
            alumno_confirmado["nombre"],
            matricula,
            identidad_confirmada=True,
            modo_contingencia=acceso_contingencia["activo"],
            persona_cache=acceso_contingencia["persona"]
        )

    cambiar_estado("Verificando conexión...", COLOR_ADVERTENCIA)
    threading.Thread(target=consultar_acceso, daemon=True).start()




def crear_pantalla_login():
    global entrada, btn_entrar, estado_var, estado_label, chk_label

    for widget in ventana.winfo_children():
        widget.destroy()

    ventana.configure(bg=COLOR_FONDO)

    main_container = tk.Frame(ventana, bg=COLOR_FONDO)
    main_container.pack(fill=tk.BOTH, expand=True)

    fondo_marca_agua = crear_fondo_marca_agua_login(
        RUTA_UTP,
        ventana.winfo_screenwidth(),
        ventana.winfo_screenheight()
    )
    if fondo_marca_agua:
        marca_agua_fondo = tk.Label(
            main_container,
            image=fondo_marca_agua,
            bg=COLOR_FONDO,
            bd=0
        )
        marca_agua_fondo.image = fondo_marca_agua
        marca_agua_fondo.place(x=0, y=0, relwidth=1, relheight=1)

    card = tk.Frame(
        main_container,
        bg=COLOR_TARJETA,
        width=TAMANOS["tarjeta_login_ancho"],
        height=TAMANOS["tarjeta_login_alto"],
        highlightthickness=1,
        highlightbackground=COLOR_BORDE
    )
    card.place(relx=0.5, rely=0.5, anchor="center")
    card.pack_propagate(False)

    if TEMA_INTERFAZ == "moderno":
        tk.Frame(card, height=5, bg=COLOR_PRIMARIO).pack(fill=tk.X)

    # LOGO
    logo = cargar_logo(RUTA_UTP, 150, 95)
    if logo:
        lbl_logo = tk.Label(card, image=logo, bg=COLOR_TARJETA)
        lbl_logo.image = logo
        lbl_logo.pack(pady=(35, 15))

    tk.Label(
        card,
        text="SISTEMA DE CONTROL\nDE LAPTOPS",
        font=("Segoe UI", 19 if TEMA_INTERFAZ == "moderno" else 18, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack()

    tk.Label(
        card,
        text="Universidad Tecnológica de Parral",
        font=("Segoe UI", 11),
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 30))

    # ENTRADA
    input_frame = tk.Frame(card, bg=COLOR_TARJETA)
    input_frame.pack(padx=50, fill=tk.X)

    tk.Label(
        input_frame,
        text="INGRESA TU MATRÍCULA",
        font=("Segoe UI", 10, "bold"),
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(anchor="w", pady=(0, 8))

    entrada = tk.Entry(
        input_frame,
        justify="center",
        **estilo_entrada()
    )
    entrada.pack(fill=tk.X, ipady=12, pady=(0, 18))

    def actualizar_timeout_desde_login(event=None):
        if entrada.get().strip() or acepta_estado_equipo.get():
            if timeout_inactividad_activo:
                reiniciar_timeout_inactividad()
            else:
                activar_timeout_inactividad()
        else:
            cancelar_timeout_inactividad()

    entrada.bind("<KeyRelease>", actualizar_timeout_desde_login, add="+")

    # FORZAR FOCO CORRECTAMENTE
    ventana.after(200, lambda: entrada.focus_set())

        # =========================
    # VALIDACIÓN SOLO NÚMEROS
    # =========================
    def solo_numeros(P):
        return P.isdigit() or P == ""

    vcmd = (ventana.register(solo_numeros), "%P")
    entrada.config(validate="key", validatecommand=vcmd)

        # =========================
        # CONFIRMACIÓN ESTADO LAPTOP
        # =========================
    global acepta_estado_equipo
    acepta_estado_equipo = tk.BooleanVar(value=False)

    check_frame = tk.Frame(card, bg=COLOR_TARJETA)
    check_frame.pack(padx=50, pady=(0, 25), fill=tk.X)

    check_content = tk.Frame(check_frame, bg=COLOR_TARJETA)
    check_content.pack(anchor="w")

    chk_label = tk.Canvas(
        check_content,
        width=32,
        height=32,
        bg=COLOR_TARJETA,
        highlightthickness=0,
        cursor="hand2"
    )
    chk_label.pack(side=tk.LEFT, padx=(0, 14))

    def dibujar_check():
        chk_label.delete("all")
        chk_label.create_rectangle(
            4, 4, 28, 28,
            outline=COLOR_PRIMARIO,
            width=3,
            fill=COLOR_TARJETA
        )
        if acepta_estado_equipo.get():
            chk_label.create_line(
                9, 17, 15, 23, 25, 10,
                fill=COLOR_EXITO,
                width=4,
                capstyle=tk.ROUND,
                joinstyle=tk.ROUND
            )

    chk_text = tk.Label(
        check_content,
        text="Confirmo que la laptop está en buen estado\n" +
             "o ya reporté cualquier anomalía.",
        font=("Segoe UI", 10),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA,
        justify="left",
        cursor="hand2",
        wraplength=300
    )
    chk_text.pack(side=tk.LEFT)

    def toggle_check(event=None):
        acepta_estado_equipo.set(not acepta_estado_equipo.get())
        actualizar_timeout_desde_login()

    acepta_estado_equipo.trace_add("write", lambda *_: dibujar_check())
    dibujar_check()
    chk_label.bind("<Button-1>", toggle_check)
    chk_text.bind("<Button-1>", toggle_check)

        # Línea divisoria
    tk.Frame(check_frame, height=1, bg=COLOR_BORDE).pack(fill=tk.X, pady=(15, 0))

  
    entrada.bind("<Return>", lambda e: iniciar_sesion())


    # =========================
    # BOTÓN INGRESAR
    # =========================
    btn_entrar = tk.Button(
        card,
        text="INGRESAR AL SISTEMA",
        padx=20,
        pady=14,
        command=iniciar_sesion,
        **estilo_boton("principal")
    )
    btn_entrar.pack(fill=tk.X, padx=50, pady=(5, 20))

    # =========================
    # ESTADO DE CONEXIÓN
    # =========================
    estado_var = tk.StringVar(value="Iniciando...")
    estado_label = tk.Label(
        card,
        textvariable=estado_var,
        font=FUENTE_PEQ_BOLD,
        fg=COLOR_EXITO,
        bg=(COLORES["info_suave"] if TEMA_INTERFAZ == "moderno" else COLOR_TARJETA),
        padx=(12 if TEMA_INTERFAZ == "moderno" else 0),
        pady=(5 if TEMA_INTERFAZ == "moderno" else 0)
    )
    estado_label.pack(pady=(5, 10))

    # =========================
    # VERSIÓN (ABAJO DEL TODO)
    # =========================
    tk.Label(
        card,
        text=f"Versión v{VERSION_SISTEMA}",
        font=FUENTE_PEQ,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(side=tk.BOTTOM, anchor="e", padx=12, pady=(5, 8))


def iniciar_conexion_en_segundo_plano():
    cambiar_estado("Verificando conexión...", COLOR_ADVERTENCIA)

    def revisar_y_conectar():
        if verificar_internet():
            conectar_google_sheets(internet_verificado=True)
        else:
            def mostrar_sin_conexion():
                # La ausencia de Internet no bloquea por sí sola una matrícula
                # que pueda validarse con la caché local de contingencia.
                cambiar_estado("Sin conexión; contingencia disponible", COLOR_ADVERTENCIA)

            ejecutar_en_ui(mostrar_sin_conexion)

    threading.Thread(target=revisar_y_conectar, daemon=True).start()

def iniciar_verificacion_periodica():
    threading.Thread(target=verificar_conexion_periodicamente, daemon=True).start()


def sincronizar_pendientes_periodicamente():
    iniciar_sincronizacion_pendientes()
    try:
        ventana.after(45000, sincronizar_pendientes_periodicamente)
    except tk.TclError:
        pass


def actualizar_configuracion_periodicamente():
    """La versión central se revisa cada diez minutos, nunca por evento."""
    threading.Thread(target=actualizar_configuracion_sincronizacion, daemon=True).start()
    try:
        ventana.after(600000, actualizar_configuracion_periodicamente)
    except tk.TclError:
        pass


def notificar_interfaz_lista():
    """Confirma al bootstrap que la ventana principal ya procesa eventos."""
    ruta_senal = os.environ.get("REGISTRO_LAPTOP_APP_LISTA")
    token_senal = os.environ.get("REGISTRO_LAPTOP_APP_TOKEN")
    if not ruta_senal or not token_senal:
        return
    ruta_temporal = f"{ruta_senal}.{os.getpid()}.tmp"
    try:
        with open(ruta_temporal, "w", encoding="ascii") as archivo:
            archivo.write(token_senal)
        os.replace(ruta_temporal, ruta_senal)
        registrar_evento_tecnico("APP_LISTA enviada")
    except OSError as error:
        registrar_evento_tecnico(f"ERROR_APP_LISTA={type(error).__name__}: {error}")
        try:
            os.remove(ruta_temporal)
        except OSError:
            pass

# --- VENTANA PRINCIPAL ---
registrar_inicio_tecnico()
registrar_evento_tecnico(
    f"MODO_PRODUCCION={not MODO_PRUEBA} "
    f"PROXIMIDAD_HABILITADA={PROXIMIDAD_HABILITADA}"
)
inicializar_infraestructura_contingencia()
ventana = tk.Tk()
vincular_dispatcher_administrativo(ventana)
ventana.bind_all("<KeyPress>", manejar_tecla_global)
ventana.bind_all("<ButtonPress>", registrar_actividad_usuario, add="+")
ventana.title(f"SISTEMA DE CONTROL DE LAPTOPS - UTP | v{VERSION_SISTEMA}")

# Pantalla completa sin bordes
ventana.attributes("-topmost", True)
ventana.state("zoomed")
ventana.overrideredirect(True)
ventana.configure(bg=COLOR_FONDO)

# AQUÍ SE CREA EL LOGIN
crear_pantalla_login()
ventana.after_idle(registrar_sesion_pendiente_al_iniciar)

# SEGURIDAD
ventana.protocol("WM_DELETE_WINDOW", cerrar_aplicacion)
ventana.bind_all("<Alt-F4>", bloquear_alt_f4)
ventana.bind_all("<Control-F4>", bloquear_alt_f4)

# --- INICIALIZACIÓN ---

# Verificar internet en segundo plano para no congelar el arranque
ventana.after(100, iniciar_limpieza_logs)
ventana.after(300, iniciar_conexion_en_segundo_plano)
ventana.after(1800, iniciar_actualizacion_en_segundo_plano)
ventana.after(5000, iniciar_verificacion_periodica)
ventana.after(15000, sincronizar_pendientes_periodicamente)
ventana.after(30000, actualizar_configuracion_periodicamente)
ventana.after(
    600,
    lambda: iniciar_heartbeat_manager(
        DIRECTORIO_APP, VERSION_SISTEMA, registrar_evento_tecnico
    )
)
ventana.after_idle(notificar_interfaz_lista)
#ventana.after(1200, mostrar_instrucciones_iniciales, "")

ventana.mainloop()
