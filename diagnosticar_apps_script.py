"""Prueba de transporte segura para el Web App de Apps Script.

No usa ni imprime la API_KEY local. Envía una clave deliberadamente inválida,
por lo que doPost debe responder NO_AUTORIZADO antes de crear Eventos o tocar
Registros. Úsese solamente desde la laptop piloto cuando esté desocupada:

    .venv\Scripts\python.exe diagnosticar_apps_script.py
"""
import json
import sys

import requests

from sincronizacion_config import cargar_configuracion_local
from sincronizador_contingencia import (
    diagnostico_respuesta_coordinador,
    url_segura,
)


def main():
    configuracion = cargar_configuracion_local(".")
    if not configuracion:
        print("CONFIGURACION_SINCRONIZACION_NO_DISPONIBLE")
        return 2

    url = configuracion["URL_APPS_SCRIPT"]
    print(f"URL={url_segura(url)}")
    payload = {
        # Esta clave es intencionalmente distinta de la API_KEY de producción.
        "api_key": "DIAGNOSTICO_SIN_AUTORIZACION",
        "id_evento": "diagnostico-sin-escritura",
        "tipo_evento": "PRESTAMO",
    }
    try:
        respuesta = requests.post(
            url, json=payload, timeout=(3, 12), allow_redirects=True,
            headers={"Accept": "application/json", "Content-Type": "application/json"},
        )
    except requests.RequestException as error:
        print(f"FALLO_RED tipo={type(error).__name__} mensaje={str(error)[:300]}")
        return 1

    print("RESPUESTA " + " ".join(
        f"{clave}={valor}" for clave, valor in
        diagnostico_respuesta_coordinador(respuesta, incluir_muestra=True).items()
    ))
    try:
        cuerpo = respuesta.json()
    except (ValueError, TypeError) as error:
        print(f"JSON_INVALIDO tipo={type(error).__name__}")
        return 1
    print("JSON=" + json.dumps(cuerpo, ensure_ascii=False)[:300])
    if respuesta.status_code == 200 and isinstance(cuerpo, dict) and cuerpo.get("codigo") == "NO_AUTORIZADO":
        print("DIAGNOSTICO_OK doPost alcanzado; no se creó ningún préstamo.")
        return 0
    print("DIAGNOSTICO_INESPERADO")
    return 1


if __name__ == "__main__":
    sys.exit(main())
