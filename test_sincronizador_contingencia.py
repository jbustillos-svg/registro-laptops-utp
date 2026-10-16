from pathlib import Path
import tempfile
import unittest

import requests

from contingencia_local import ContingenciaLocal
from sincronizador_contingencia import SincronizadorContingencia, enviar_evento_coordinador


CONFIG = {
    "URL_APPS_SCRIPT": "https://script.google.com/macros/s/ejemplo/exec",
    "API_KEY": "secreto-prueba",
    "VERSION_CONFIG": "1",
}


class RespuestaFalsa:
    status_code = 200

    def __init__(self, cuerpo):
        self.cuerpo = cuerpo

    def json(self):
        return self.cuerpo


class PruebasClienteCoordinador(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.almacen = ContingenciaLocal(Path(self.temporal.name))
        self.almacen.inicializar()
        self.logs = []
        self.evento = self.almacen.crear_evento(
            "LAP-1", "PRESTAMO", "A1", "Ana", "2026-10-16", "10:00:00",
            "LAP-1", {"fecha": "2026-10-16", "hora_ingreso": "10:00:00", "nombre": "Ana"},
            id_evento="12345678-1234-1234-1234-123456789012",
        )

    def tearDown(self):
        self.temporal.cleanup()

    def test_post_contiene_evento_y_nunca_se_registra_clave(self):
        solicitudes = []

        def post(url, json, timeout):
            solicitudes.append((url, json, timeout))
            return RespuestaFalsa({"estado": "APLICADO", "referencia_registro": "Registros!2"})

        cantidad = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append, post=post
        ).sincronizar()
        self.assertEqual(cantidad, 1)
        self.assertEqual(solicitudes[0][1]["id_evento"], self.evento)
        self.assertEqual(solicitudes[0][1]["api_key"], CONFIG["API_KEY"])
        self.assertFalse(any(CONFIG["API_KEY"] in log for log in self.logs))
        self.assertEqual(self.almacen.obtener_eventos_pendientes(), [])

    def test_timeout_deja_pendiente_y_reintento_ya_aplicado_confirma(self):
        def timeout(*_args, **_kwargs):
            raise requests.Timeout("sin respuesta")

        cliente = SincronizadorContingencia(self.almacen, CONFIG, self.logs.append, post=timeout)
        self.assertEqual(cliente.sincronizar(), 0)
        self.assertEqual(self.almacen.obtener_evento(self.evento)["intentos"], 1)

        cliente = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda *_a, **_k: RespuestaFalsa({"estado": "YA_APLICADO", "referencia_registro": "Registros!2"}),
        )
        self.assertEqual(cliente.sincronizar(), 1)
        self.assertEqual(self.almacen.obtener_eventos_pendientes(), [])

    def test_sin_configuracion_no_intenta_post_y_deja_pendiente(self):
        llamadas = []
        cliente = SincronizadorContingencia(
            self.almacen, {}, self.logs.append,
            post=lambda *_a, **_k: llamadas.append(True),
        )
        self.assertEqual(cliente.sincronizar(), 0)
        self.assertEqual(llamadas, [])
        self.assertEqual(len(self.almacen.obtener_eventos_pendientes()), 1)

    def test_sheets_caido_con_cache_endpoint_disponible_sin_dependencia_gspread(self):
        cliente = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda *_a, **_k: RespuestaFalsa({"estado": "APLICADO"}),
        )
        self.assertEqual(cliente.sincronizar(), 1)

    def test_envio_inmediato_y_reintento_usan_mismo_uuid(self):
        evento = self.almacen.obtener_evento(self.evento)
        enviados = []
        def post(_url, json, timeout):
            enviados.append(json["id_evento"])
            return RespuestaFalsa({"estado": "YA_APLICADO"})
        enviar_evento_coordinador(evento, CONFIG, post=post)
        enviar_evento_coordinador(evento, CONFIG, post=post)
        self.assertEqual(enviados, [self.evento, self.evento])


if __name__ == "__main__":
    unittest.main()
