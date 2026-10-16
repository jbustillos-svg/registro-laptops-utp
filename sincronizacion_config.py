"""Aprovisionamiento seguro de la configuración del coordinador.

La clave no se registra ni se versiona. La hoja central se consulta sólo al
conectar y de forma periódica; el sincronizador usa siempre esta copia local.
"""
import json
import os
from pathlib import Path


ARCHIVO_CONFIGURACION_SINCRONIZACION = "sincronizacion_config.json"
CLAVES_REQUERIDAS = ("URL_APPS_SCRIPT", "API_KEY", "VERSION_CONFIG")


def _ruta(directorio_app):
    return Path(directorio_app) / ARCHIVO_CONFIGURACION_SINCRONIZACION


def validar_configuracion(configuracion):
    """Devuelve una copia normalizada o ``None``; nunca revela valores."""
    if not isinstance(configuracion, dict):
        return None
    resultado = {}
    for clave in CLAVES_REQUERIDAS:
        valor = configuracion.get(clave)
        if not isinstance(valor, (str, int)) or not str(valor).strip():
            return None
        resultado[clave] = str(valor).strip()
    if not resultado["URL_APPS_SCRIPT"].startswith("https://script.google.com/"):
        return None
    return resultado


def cargar_configuracion_local(directorio_app):
    try:
        with _ruta(directorio_app).open("r", encoding="utf-8") as archivo:
            return validar_configuracion(json.load(archivo))
    except (OSError, ValueError, TypeError):
        return None


def guardar_configuracion_local(directorio_app, configuracion):
    """Reemplazo atómico sólo de una configuración previamente validada."""
    configuracion = validar_configuracion(configuracion)
    if configuracion is None:
        return False
    ruta = _ruta(directorio_app)
    temporal = ruta.with_suffix(ruta.suffix + ".tmp")
    try:
        ruta.parent.mkdir(parents=True, exist_ok=True)
        with temporal.open("w", encoding="utf-8") as archivo:
            json.dump(configuracion, archivo, ensure_ascii=False, sort_keys=True)
            archivo.flush()
            os.fsync(archivo.fileno())
        os.replace(str(temporal), str(ruta))
        return True
    except OSError:
        try:
            temporal.unlink(missing_ok=True)
        except OSError:
            pass
        return False


def leer_configuracion_hoja(hoja):
    """Extrae CLAVE/VALOR sin crear ni modificar hojas de producción."""
    filas = hoja.get_all_values()
    if not filas:
        return None
    encabezados = [str(valor).strip().upper() for valor in filas[0]]
    try:
        indice_clave = encabezados.index("CLAVE")
        indice_valor = encabezados.index("VALOR")
    except ValueError:
        return None
    valores = {}
    for fila in filas[1:]:
        if len(fila) <= max(indice_clave, indice_valor):
            continue
        clave = str(fila[indice_clave]).strip().upper()
        if clave in CLAVES_REQUERIDAS:
            valores[clave] = str(fila[indice_valor]).strip()
    return validar_configuracion(valores)


def actualizar_configuracion_desde_hoja(directorio_app, hoja):
    """Devuelve la configuración vigente y si se actualizó el caché local."""
    remota = leer_configuracion_hoja(hoja)
    local = cargar_configuracion_local(directorio_app)
    if remota is None:
        return local, False
    if remota == local:
        return local, False
    if guardar_configuracion_local(directorio_app, remota):
        return remota, True
    return local, False
