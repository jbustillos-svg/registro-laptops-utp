from pathlib import Path
import json
import tempfile
import unittest

import requests

from contingencia_local import ContingenciaLocal
from sincronizador_contingencia import (
    ErrorCoordinador,
    SincronizadorContingencia,
    diagnostico_error_coordinador,
    enviar_evento_coordinador,
)


CONFIG = {
    "URL_APPS_SCRIPT": "https://script.google.com/macros/s/ejemplo/exec",
    "API_KEY": "secreto-prueba",
    "VERSION_CONFIG": "1",
}


class RespuestaFalsa:
    status_code = 200

    def __init__(self, cuerpo, texto="", status_code=200, headers=None, history=None):
        self.cuerpo = cuerpo
        self.text = texto
        self.status_code = status_code
        self.headers = headers or {"Content-Type": "application/json"}
        self.history = history or []
        self.url = CONFIG["URL_APPS_SCRIPT"]

    def json(self):
        if isinstance(self.cuerpo, Exception):
            raise self.cuerpo
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

        def post(url, json, timeout, **kwargs):
            solicitudes.append((url, json, timeout, kwargs))
            return RespuestaFalsa({"estado": "APLICADO", "referencia_registro": "Registros!2"})

        cantidad = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append, post=post
        ).sincronizar()
        self.assertEqual(cantidad, 1)
        self.assertEqual(solicitudes[0][1]["id_evento"], self.evento)
        self.assertEqual(solicitudes[0][1]["api_key"], CONFIG["API_KEY"])
        self.assertTrue(solicitudes[0][3]["allow_redirects"])
        self.assertEqual(solicitudes[0][3]["headers"]["Content-Type"], "application/json")
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

    def test_evento_nuevo_no_es_bloqueado_por_error_historico(self):
        self.almacen.marcar_sincronizado(self.evento)
        viejo = self.almacen.crear_evento(
            "LAP-0", "CIERRE_AUTOMATICO", "A0", "Anterior", "2026-10-01", "08:00:00",
            "LAP-0", {}, id_evento="viejo-error", fecha_creacion="2026-10-01T08:00:00"
        )
        nuevo = self.almacen.crear_evento(
            "LAP-1", "PRESTAMO", "A1", "Ana", "2026-10-21", "10:00:00",
            "LAP-1", {}, id_evento="nuevo-actual", fecha_creacion="2026-10-21T10:00:00"
        )
        enviados = []

        def post(_url, json, timeout, **_k):
            enviados.append(json["id_evento"])
            if json["id_evento"] == viejo:
                return RespuestaFalsa({"codigo": "DEPENDENCIA_PRESTAMO_NO_ENCONTRADA"})
            return RespuestaFalsa({"estado": "APLICADO"})

        cliente = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append, post=post,
        )

        self.assertEqual(cliente.sincronizar(), 1)
        self.assertEqual(enviados, [nuevo, viejo])
        self.assertEqual(self.almacen.obtener_evento(nuevo)["sincronizado"], 1)
        self.assertEqual(self.almacen.obtener_evento(viejo)["sincronizado"], 0)
        self.assertTrue(any("SINCRONIZACION_EVENTO_CONTINUA_TRAS_ERROR" in log for log in self.logs))

    def test_devolucion_se_envia_despues_del_prestamo_de_la_misma_sesion(self):
        self.almacen.marcar_sincronizado(self.evento)
        sesion = self.almacen.registrar_prestamo_offline(
            "A2", "Beto", "LAP-2", "2026-10-21", "09:00:00"
        )
        entrega = self.almacen.registrar_devolucion_offline(
            "A2", "LAP-2", "2026-10-21", "12:00:00", id_sesion=sesion["id_sesion"]
        )
        enviados = []
        cantidad = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda _url, json, timeout, **_k: (
                enviados.append(json["id_evento"])
                or RespuestaFalsa({"estado": "APLICADO"})
            ),
        ).sincronizar()

        self.assertEqual(cantidad, 2)
        self.assertEqual(enviados, [sesion["prestamo_evento_id"], entrega["devolucion_evento_id"]])

    def test_devolucion_no_se_envia_si_su_prestamo_pendiente_falla(self):
        self.almacen.marcar_sincronizado(self.evento)
        sesion = self.almacen.registrar_prestamo_offline(
            "A3", "Caro", "LAP-3", "2026-10-21", "09:00:00"
        )
        entrega = self.almacen.registrar_devolucion_offline(
            "A3", "LAP-3", "2026-10-21", "12:00:00", id_sesion=sesion["id_sesion"]
        )
        enviados = []
        cantidad = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda _url, json, timeout, **_k: (
                enviados.append(json["id_evento"])
                or (_ for _ in ()).throw(requests.Timeout())
            ),
        ).sincronizar()

        self.assertEqual(cantidad, 0)
        self.assertEqual(enviados, [sesion["prestamo_evento_id"]])
        self.assertEqual(self.almacen.obtener_evento(entrega["devolucion_evento_id"])["sincronizado"], 0)
        self.assertTrue(any("DEPENDENCIA_PRESTAMO_PENDIENTE" in log for log in self.logs))

    def test_error_historico_se_difiere_en_el_siguiente_ciclo(self):
        self.almacen.marcar_sincronizado(self.evento)
        viejo = self.almacen.crear_evento(
            "LAP-0", "CIERRE_AUTOMATICO", "A0", "Anterior", "2026-10-01", "08:00:00",
            "LAP-0", {}, id_evento="viejo-backoff", fecha_creacion="2026-10-01T08:00:00"
        )
        self.almacen.incrementar_intentos(viejo, "DEPENDENCIA_PRESTAMO_NO_ENCONTRADA")
        enviados = []
        cliente = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda _url, json, timeout, **_k: enviados.append(json["id_evento"]),
        )

        self.assertEqual(cliente.sincronizar(), 0)
        self.assertEqual(enviados, [])
        self.assertTrue(any("id=viejo-backoff motivo=BACKOFF" in log for log in self.logs))

    def test_cada_uuid_pendiente_se_envia_una_sola_vez_por_ciclo(self):
        self.almacen.marcar_sincronizado(self.evento)
        primero = self.almacen.crear_evento(
            "LAP-4", "PRESTAMO", "A4", "Dani", "2026-10-21", "09:00:00",
            "LAP-4", {}, id_evento="uuid-unico-1"
        )
        segundo = self.almacen.crear_evento(
            "LAP-5", "PRESTAMO", "A5", "Eva", "2026-10-21", "09:01:00",
            "LAP-5", {}, id_evento="uuid-unico-2"
        )
        enviados = []
        SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda _url, json, timeout, **_k: (
                enviados.append(json["id_evento"])
                or RespuestaFalsa({"estado": "APLICADO"})
            ),
        ).sincronizar()

        self.assertEqual(enviados, [primero, segundo])
        self.assertEqual(len(enviados), len(set(enviados)))

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
        def post(_url, json, timeout, **_kwargs):
            enviados.append(json["id_evento"])
            return RespuestaFalsa({"estado": "YA_APLICADO"})
        enviar_evento_coordinador(evento, CONFIG, post=post)
        enviar_evento_coordinador(evento, CONFIG, post=post)
        self.assertEqual(enviados, [self.evento, self.evento])

    def test_cada_evento_critico_persistido_permanece_pendiente_si_falla_el_worker(self):
        """El disparador inmediato puede fallar; SQLite conserva el mismo UUID."""
        self.almacen.marcar_sincronizado(self.evento)
        for indice, tipo in enumerate(
            ("PRESTAMO", "DEVOLUCION", "NO_ENTREGA", "CIERRE_AUTOMATICO"), start=1
        ):
            with self.subTest(tipo=tipo):
                id_evento = f"00000000-0000-4000-8000-{indice:012d}"
                creado = self.almacen.crear_evento(
                    "LAP-1", tipo, "A1", "Ana", "2026-10-16", "10:00:00",
                    "LAP-1", {"id_sesion": "sesion-prueba"}, id_evento=id_evento,
                )

                cliente = SincronizadorContingencia(
                    self.almacen, CONFIG, self.logs.append,
                    post=lambda *_a, **_k: (_ for _ in ()).throw(requests.Timeout()),
                )
                self.assertEqual(cliente.sincronizar(), 0)
                pendiente = self.almacen.obtener_evento(creado)
                self.assertEqual((pendiente["id_evento"], pendiente["tipo_evento"]), (id_evento, tipo))
                self.assertEqual(pendiente["sincronizado"], 0)

                enviados = []
                cliente = SincronizadorContingencia(
                    self.almacen, CONFIG, self.logs.append,
                    post=lambda _url, json, timeout, **_k: (
                        enviados.append(json["id_evento"])
                        or RespuestaFalsa({"estado": "APLICADO"})
                    ),
                )
                self.assertEqual(cliente.sincronizar(), 1)
                self.assertEqual(enviados, [id_evento])
                self.assertEqual(self.almacen.obtener_eventos_pendientes(), [])

    def test_devolucion_pre_apagado_rapida_confirma_hora_original(self):
        sesion = self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        entrega = self.almacen.registrar_devolucion_offline(
            "A1", "LAP-1", "2026-10-16", "14:00:17",
            id_sesion=sesion["id_sesion"],
        )
        enviados = []
        cliente = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda _url, json, timeout, **_k: (
                enviados.append((json, timeout))
                or RespuestaFalsa({"estado": "APLICADO", "referencia_registro": "Registros!3"})
            ),
        )

        self.assertTrue(cliente.sincronizar_evento(entrega["devolucion_evento_id"]))
        self.assertEqual(enviados[0][0]["hora_evento"], "14:00:17")
        self.assertEqual(enviados[0][1], (2, 3))
        self.assertEqual(
            self.almacen.obtener_evento(entrega["devolucion_evento_id"])["hora_evento"],
            "14:00:17",
        )
        self.assertEqual(
            self.almacen.obtener_evento(entrega["devolucion_evento_id"])["sincronizado"], 1
        )

    def test_devolucion_pre_apagado_sin_internet_queda_pendiente(self):
        # Se crea una sesión explícita para que la devolución sea durable.
        sesion = self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        entrega = self.almacen.registrar_devolucion_offline(
            "A1", "LAP-1", "2026-10-16", "14:00:17",
            id_sesion=sesion["id_sesion"],
        )
        cliente = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda *_a, **_k: (_ for _ in ()).throw(requests.ConnectionError()),
        )

        self.assertFalse(cliente.sincronizar_evento(entrega["devolucion_evento_id"]))
        pendiente = self.almacen.obtener_evento(entrega["devolucion_evento_id"])
        self.assertEqual((pendiente["sincronizado"], pendiente["hora_evento"]), (0, "14:00:17"))

        enviados = []
        reintento = SincronizadorContingencia(
            self.almacen, CONFIG, self.logs.append,
            post=lambda _url, json, timeout, **_k: (
                enviados.append(json["hora_evento"])
                or RespuestaFalsa({"estado": "APLICADO"})
            ),
        )
        self.assertTrue(reintento.sincronizar_evento(entrega["devolucion_evento_id"]))
        self.assertEqual(enviados, ["14:00:17"])

    def test_devolucion_pre_apagado_timeout_o_http_deja_evento_pendiente(self):
        for causa in (
            lambda: (_ for _ in ()).throw(requests.Timeout()),
            lambda: RespuestaFalsa({}, texto="error", status_code=503),
        ):
            with self.subTest(causa=causa):
                sesion = self.almacen.registrar_prestamo_offline(
                    "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
                )
                entrega = self.almacen.registrar_devolucion_offline(
                    "A1", "LAP-1", "2026-10-16", "14:00:17",
                    id_sesion=sesion["id_sesion"],
                )
                cliente = SincronizadorContingencia(
                    self.almacen, CONFIG, self.logs.append,
                    post=lambda *_a, **_k: causa(),
                )
                self.assertFalse(cliente.sincronizar_evento(entrega["devolucion_evento_id"]))
                self.assertEqual(
                    self.almacen.obtener_evento(entrega["devolucion_evento_id"])["sincronizado"], 0
                )

    def test_flujo_durable_preserva_matricula_y_hora_de_diana_de_extremo_python(self):
        """SQLite y payload conservan el identificador literal antes del Apps Script."""
        matricula = "0824118047"
        nombre = "Diana Chaparro Leyva"
        laptop = "AulaMovil24-B"
        sesion = self.almacen.registrar_prestamo_coordinador(
            matricula, nombre, laptop, "2026-10-19", "14:05:45"
        )
        prestamo = self.almacen.obtener_evento(sesion["prestamo_evento_id"])
        self.assertEqual(sesion["matricula"], matricula)
        self.assertEqual(prestamo["matricula"], matricula)
        self.assertEqual(
            SincronizadorContingencia._payload(prestamo, CONFIG)["matricula"], matricula
        )
        self.assertTrue(
            SincronizadorContingencia(
                self.almacen, CONFIG, self.logs.append,
                post=lambda *_a, **_k: RespuestaFalsa(
                    {"estado": "APLICADO", "referencia_registro": "Registros!9"}
                ),
            ).sincronizar_evento(sesion["prestamo_evento_id"])
        )

        entrega = self.almacen.registrar_devolucion_offline(
            matricula, laptop, "2026-10-19", "14:06:23", id_sesion=sesion["id_sesion"]
        )
        devolucion = self.almacen.obtener_evento(entrega["devolucion_evento_id"])
        datos = json.loads(devolucion["datos_json"])
        self.assertEqual((devolucion["matricula"], datos["hora_salida"]), (matricula, "14:06:23"))
        self.assertEqual(datos["prestamo_evento_id"], sesion["prestamo_evento_id"])
        enviado = []
        self.assertTrue(
            SincronizadorContingencia(
                self.almacen, CONFIG, self.logs.append,
                post=lambda _url, json, timeout, **_k: (
                    enviado.append(json)
                    or RespuestaFalsa({"estado": "APLICADO", "referencia_registro": "Registros!9"})
                ),
            ).sincronizar_evento(entrega["devolucion_evento_id"]),
            self.logs,
        )
        self.assertEqual(enviado[0]["matricula"], matricula)
        self.assertEqual(enviado[0]["datos"]["hora_salida"], "14:06:23")

    def test_html_con_http_200_conserva_detalle_tecnico_sin_clave(self):
        evento = self.almacen.obtener_evento(self.evento)
        respuesta = RespuestaFalsa(
            ValueError("JSON inválido"), texto="<html>portal api_key=secreto-prueba</html>",
            headers={"Content-Type": "text/html"},
        )
        with self.assertRaises(ErrorCoordinador) as contexto:
            enviar_evento_coordinador(evento, CONFIG, post=lambda *_a, **_k: respuesta)
        detalle = diagnostico_error_coordinador(contexto.exception)
        self.assertIn("RESPUESTA_COORDINADOR_INVALIDA", detalle)
        self.assertIn("content_type=text/html", detalle)
        self.assertIn("muestra=<html>portal api_key=[REDACTADO]", detalle)
        self.assertNotIn(CONFIG["API_KEY"], detalle)

    def test_http_y_redirect_se_registran_en_error_tecnico(self):
        previo = RespuestaFalsa(None, status_code=302, headers={
            "Location": "https://script.googleusercontent.com/macros/resultado?token=privado"
        })
        previo.url = CONFIG["URL_APPS_SCRIPT"]
        previo.request = type("Solicitud", (), {"method": "POST"})()
        respuesta = RespuestaFalsa(None, texto="error api_key=secreto-prueba", status_code=502,
                                  headers={"Content-Type": "text/html"}, history=[previo])
        respuesta.url = "https://script.googleusercontent.com/macros/resultado?token=privado"
        respuesta.request = type("Solicitud", (), {"method": "GET"})()
        with self.assertRaises(ErrorCoordinador) as contexto:
            enviar_evento_coordinador(
                self.almacen.obtener_evento(self.evento), CONFIG,
                post=lambda *_a, **_k: respuesta,
            )
        detalle = diagnostico_error_coordinador(contexto.exception)
        self.assertIn("HTTP_502", detalle)
        self.assertIn("302:POST:https://script.google.com/macros/s/ejemplo/exec", detalle)
        self.assertIn("final:GET:https://script.googleusercontent.com/macros/resultado", detalle)
        self.assertNotIn("token=privado", detalle)
        self.assertNotIn(CONFIG["API_KEY"], detalle)


if __name__ == "__main__":
    unittest.main()
