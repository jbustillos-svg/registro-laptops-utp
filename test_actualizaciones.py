import importlib.machinery
import importlib.util
from pathlib import Path
import subprocess
import tempfile
import time
import unittest
from unittest import mock


def cargar_bootstrap():
    ruta = Path(__file__).with_name("bootstrap_laptop.pyw")
    loader = importlib.machinery.SourceFileLoader("bootstrap_pruebas", str(ruta))
    spec = importlib.util.spec_from_loader(loader.name, loader)
    modulo = importlib.util.module_from_spec(spec)
    loader.exec_module(modulo)
    return modulo


bootstrap = cargar_bootstrap()


class AvisoFalso:
    def __init__(self):
        self.mensajes = []

    def actualizar(self, mensaje):
        self.mensajes.append(mensaje)


def git(ruta, *argumentos):
    return subprocess.run(
        ["git", *argumentos], cwd=ruta, check=True, capture_output=True, text=True
    )


class PruebasActualizacion(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        raiz = Path(self.temporal.name)
        self.remoto = raiz / "remoto.git"
        self.autor = raiz / "autor"
        self.laptop = raiz / "laptop"
        git(raiz, "init", "--bare", str(self.remoto))
        git(raiz, "clone", str(self.remoto), str(self.autor))
        git(self.autor, "config", "user.email", "pruebas@example.com")
        git(self.autor, "config", "user.name", "Pruebas")
        git(self.autor, "switch", "-c", "main")
        (self.autor / "app.txt").write_text("v1\n", encoding="utf-8")
        git(self.autor, "add", "app.txt")
        git(self.autor, "commit", "-m", "v1")
        git(self.autor, "push", "-u", "origin", "main")
        git(raiz, "clone", "--branch", "main", str(self.remoto), str(self.laptop))
        git(self.laptop, "config", "user.email", "laptop@example.com")
        git(self.laptop, "config", "user.name", "Laptop")
        self._publicar_v2()

        self.parches = [
            mock.patch.object(bootstrap, "DIRECTORIO_APP", self.laptop),
            mock.patch.object(
                bootstrap,
                "RUTA_ACTUALIZACION_PENDIENTE",
                self.laptop / ".actualizacion_pendiente",
            ),
            mock.patch.object(
                bootstrap, "RUTA_BLOQUEO_ACTUALIZACION", self.laptop / ".actualizacion.lock"
            ),
            mock.patch.object(bootstrap, "RUTA_LOG", self.laptop / "bootstrap.log"),
            mock.patch.object(bootstrap, "RUTA_LOG_ARRANQUE", self.laptop / "arranque.log"),
            mock.patch.object(bootstrap, "red_disponible", return_value=True),
        ]
        for parche in self.parches:
            parche.start()

    def tearDown(self):
        for parche in reversed(self.parches):
            parche.stop()
        self.temporal.cleanup()

    def _publicar_v2(self):
        (self.autor / "app.txt").write_text("v2\n", encoding="utf-8")
        git(self.autor, "commit", "-am", "v2")
        git(self.autor, "push")

    def _marcar(self):
        bootstrap.RUTA_ACTUALIZACION_PENDIENTE.write_text(
            "pendiente\n", encoding="ascii"
        )

    def test_sin_marker_no_hace_operacion_remota(self):
        with mock.patch.object(bootstrap, "ejecutar_git") as ejecutar_git:
            self.assertTrue(bootstrap.aplicar_actualizacion_pendiente(AvisoFalso()))
            comandos = [llamada.args[0][0] for llamada in ejecutar_git.call_args_list]
            self.assertNotIn("fetch", comandos)
            self.assertNotIn("merge", comandos)

    def test_fast_forward_aplica_y_elimina_marker(self):
        self._marcar()
        self.assertTrue(bootstrap.aplicar_actualizacion_pendiente(AvisoFalso()))
        self.assertEqual((self.laptop / "app.txt").read_text(), "v2\n")
        self.assertFalse(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())

    def test_sin_marker_detecta_y_aplica_en_mismo_arranque(self):
        aviso = AvisoFalso()
        self.assertTrue(bootstrap.comprobar_actualizacion_arranque(aviso))
        self.assertEqual((self.laptop / "app.txt").read_text(), "v2\n")
        self.assertFalse(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())
        self.assertIn("Actualizando sistema...", aviso.mensajes)

    def test_igual_a_origin_arranca_sin_marker(self):
        git(self.laptop, "pull", "--ff-only")
        inicio = time.monotonic()
        self.assertTrue(bootstrap.comprobar_actualizacion_arranque(AvisoFalso()))
        self.assertFalse(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())
        self.assertLess(time.monotonic() - inicio, 3)

    def test_sin_marker_y_remoto_inaccesible_usa_version_local(self):
        fallo = subprocess.CompletedProcess(["git", "fetch"], 1, "", "sin red")
        with mock.patch.object(bootstrap, "ejecutar_git", return_value=fallo) as ejecutar:
            self.assertTrue(bootstrap.comprobar_actualizacion_arranque(AvisoFalso()))
        fetches = [
            llamada for llamada in ejecutar.call_args_list
            if llamada.args[0][0] == "fetch"
        ]
        self.assertEqual(len(fetches), bootstrap.INTENTOS_FETCH_RAPIDO)
        self.assertFalse(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())

    def test_untracked_no_bloquea(self):
        (self.laptop / "local.txt").write_text("local\n", encoding="utf-8")
        self._marcar()
        self.assertTrue(bootstrap.aplicar_actualizacion_pendiente(AvisoFalso()))
        self.assertTrue((self.laptop / "local.txt").exists())

    def test_tracked_bloquea_y_conserva_marker(self):
        (self.laptop / "app.txt").write_text("cambio local\n", encoding="utf-8")
        self._marcar()
        self.assertFalse(bootstrap.aplicar_actualizacion_pendiente(AvisoFalso()))
        self.assertTrue(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())
        self.assertEqual((self.laptop / "app.txt").read_text(), "cambio local\n")

    def test_no_fast_forward_bloquea_y_conserva_marker(self):
        (self.laptop / "local.txt").write_text("commit local\n", encoding="utf-8")
        git(self.laptop, "add", "local.txt")
        git(self.laptop, "commit", "-m", "rama local")
        self._marcar()
        self.assertFalse(bootstrap.aplicar_actualizacion_pendiente(AvisoFalso()))
        self.assertTrue(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())

    def test_estado_legado_crea_marker(self):
        git(self.laptop, "fetch", "origin", "main")
        bootstrap.adoptar_actualizacion_legada()
        self.assertTrue(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())

    def test_fallo_fetch_conserva_marker(self):
        self._marcar()
        original = bootstrap.ejecutar_git

        def ejecutar(argumentos, timeout=20):
            if argumentos[0] == "fetch":
                return subprocess.CompletedProcess(argumentos, 1, "", "sin red")
            return original(argumentos, timeout)

        with mock.patch.object(bootstrap, "ejecutar_git", side_effect=ejecutar):
            self.assertFalse(bootstrap.aplicar_actualizacion_pendiente(AvisoFalso()))
        self.assertTrue(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())

    def test_sin_red_espera_y_luego_aplica(self):
        self._marcar()
        aviso = AvisoFalso()
        with mock.patch.object(bootstrap, "red_disponible", return_value=False), mock.patch.object(
            bootstrap, "esperar_red"
        ) as esperar:
            self.assertTrue(bootstrap.aplicar_actualizacion_pendiente(aviso))
        esperar.assert_called_once_with(aviso)
        self.assertFalse(bootstrap.RUTA_ACTUALIZACION_PENDIENTE.exists())

    def test_marker_tiene_prioridad_sobre_comprobacion_rapida(self):
        self._marcar()
        with mock.patch.object(bootstrap, "aplicar_actualizacion_pendiente", return_value=True) as aplicar:
            self.assertTrue(bootstrap.comprobar_actualizacion_arranque(AvisoFalso()))
        aplicar.assert_called_once()

    def test_hash_requirements_solo_invalida_si_cambia(self):
        requirements = self.laptop / "requirements..txt"
        marcador = self.laptop / ".requirements.sha256"
        python = self.laptop / "python.exe"
        pythonw = self.laptop / "pythonw.exe"
        requirements.write_text("requests==1\n", encoding="utf-8")
        python.touch()
        pythonw.touch()
        with mock.patch.object(bootstrap, "RUTA_REQUIREMENTS", requirements), mock.patch.object(
            bootstrap, "RUTA_MARCADOR", marcador
        ), mock.patch.object(bootstrap, "RUTA_PYTHON", python), mock.patch.object(
            bootstrap, "RUTA_PYTHONW", pythonw
        ), mock.patch.object(bootstrap, "verificar_imports_entorno", return_value=True):
            marcador.write_text(bootstrap.hash_requirements() + "\n", encoding="ascii")
            self.assertTrue(bootstrap.entorno_preparado(bootstrap.hash_requirements()))
            requirements.write_text("requests==2\n", encoding="utf-8")
            self.assertFalse(bootstrap.entorno_preparado(bootstrap.hash_requirements()))

    def test_app_lista_conserva_contrato(self):
        señal = self.laptop / "app_lista.tmp"
        señal.write_text("token", encoding="ascii")

        class Proceso:
            returncode = None

            def poll(self):
                return None

        self.assertTrue(bootstrap.esperar_app_lista(Proceso(), señal, "token", timeout=0.2))

    def test_solo_bootstrap_contiene_merge(self):
        raiz = Path(__file__).parent
        self.assertIn('"merge", "--ff-only"', (raiz / "bootstrap_laptop.pyw").read_text(encoding="utf-8"))
        self.assertNotIn('"merge"', (raiz / "actualizador_segundo_plano.pyw").read_text(encoding="utf-8"))
        self.assertNotIn('"merge"', (raiz / "registro_laptop.pyw").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
