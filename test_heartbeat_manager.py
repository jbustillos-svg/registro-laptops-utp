import json
import socket
import tempfile
import threading
import time
import unittest
import urllib.error
from pathlib import Path
from unittest import mock

import heartbeat_manager


class Respuesta:
    def __init__(self, estado=200):
        self.estado = estado

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def getcode(self):
        return self.estado


class HeartbeatManagerTests(unittest.TestCase):
    def setUp(self):
        heartbeat_manager._worker = None
        self.temporal = tempfile.TemporaryDirectory()
        self.directorio = Path(self.temporal.name)

    def tearDown(self):
        worker = heartbeat_manager._worker
        if worker is not None:
            worker._detener.set()
            worker.thread.join(timeout=0.2)
        heartbeat_manager._worker = None
        self.temporal.cleanup()

    def escribir_config(self, habilitado=True):
        config = {
            "habilitado": habilitado,
            "url": "http://manager.local:5000",
            "api_key": "secreto",
        }
        (self.directorio / "manager_config.json").write_text(
            json.dumps(config), encoding="utf-8"
        )

    def test_config_inexistente_malformada_y_deshabilitada(self):
        registrar = mock.Mock()
        self.assertIsNone(heartbeat_manager.iniciar_heartbeat_manager(
            str(self.directorio), "1.3.2", registrar
        ))
        (self.directorio / "manager_config.json").write_text("{", encoding="utf-8")
        self.assertIsNone(heartbeat_manager.iniciar_heartbeat_manager(
            str(self.directorio), "1.3.2", registrar
        ))
        self.escribir_config(habilitado=False)
        self.assertIsNone(heartbeat_manager.iniciar_heartbeat_manager(
            str(self.directorio), "1.3.2", registrar
        ))
        registrar.assert_not_called()

    def test_envia_hostname_version_header_y_endpoint(self):
        capturado = {}

        def urlopen(solicitud, timeout):
            capturado["solicitud"] = solicitud
            capturado["timeout"] = timeout
            return Respuesta(200)

        eventos = []
        worker = heartbeat_manager.HeartbeatManager(
            {"url": "http://manager.local:5000", "api_key": "secreto"},
            "1.3.2",
            eventos.append,
            hostname=socket.gethostname(),
            urlopen=urlopen,
        )
        worker._enviar()
        solicitud = capturado["solicitud"]
        self.assertEqual(solicitud.full_url, "http://manager.local:5000/api/heartbeat")
        self.assertEqual(solicitud.method, "POST")
        self.assertEqual(capturado["timeout"], 3)
        self.assertEqual(solicitud.headers["X-laptop-manager-key"], "secreto")
        self.assertEqual(
            json.loads(solicitud.data),
            {"equipo": socket.gethostname(), "version": "1.3.2"},
        )
        self.assertIn("HEARTBEAT_MANAGER enviado", eventos)

    def test_manager_no_disponible_timeout_y_401_no_escapan(self):
        casos = [
            OSError("sin red"),
            TimeoutError("timeout"),
            urllib.error.HTTPError("url", 401, "Unauthorized", {}, None),
        ]
        for error in casos:
            eventos = []

            def fallar(_solicitud, timeout, error=error):
                raise error

            worker = heartbeat_manager.HeartbeatManager(
                {"url": "http://manager.local:5000", "api_key": "secreto"},
                "1.3.2",
                eventos.append,
                urlopen=fallar,
            )
            worker._enviar()
            self.assertTrue(eventos[-1].startswith(
                "HEARTBEAT_MANAGER no disponible"
            ))

    def test_inicio_no_bloquea_no_duplica_y_thread_es_daemon(self):
        self.escribir_config()
        bloqueo = threading.Event()

        def lento(_solicitud, timeout):
            bloqueo.wait(0.15)
            return Respuesta(200)

        with mock.patch("heartbeat_manager.urllib.request.urlopen", lento):
            inicio = time.monotonic()
            primero = heartbeat_manager.iniciar_heartbeat_manager(
                str(self.directorio), "1.3.2", mock.Mock(), intervalo=60
            )
            segundo = heartbeat_manager.iniciar_heartbeat_manager(
                str(self.directorio), "1.3.2", mock.Mock(), intervalo=60
            )
            duracion = time.monotonic() - inicio
            self.assertIs(primero, segundo)
            self.assertLess(duracion, 0.1)
            self.assertTrue(primero.thread.daemon)
            bloqueo.set()


if __name__ == "__main__":
    unittest.main()
