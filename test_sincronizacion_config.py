import json
from pathlib import Path
import subprocess
import tempfile
import unittest

from sincronizacion_config import (
    actualizar_configuracion_desde_hoja,
    cargar_configuracion_local,
    guardar_configuracion_local,
)


CONFIG = {
    "URL_APPS_SCRIPT": "https://script.google.com/macros/s/ejemplo/exec",
    "API_KEY": "secreto-prueba",
    "VERSION_CONFIG": "1",
}


class HojaFalsa:
    def __init__(self, filas):
        self.filas = filas

    def get_all_values(self):
        return self.filas


class PruebasConfiguracionSincronizacion(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.directorio = Path(self.temporal.name)

    def tearDown(self):
        self.temporal.cleanup()

    @staticmethod
    def hoja(configuracion):
        return HojaFalsa([
            ["CLAVE", "VALOR"],
            *[[clave, valor] for clave, valor in configuracion.items()],
        ])

    def test_descarga_y_cachea_configuracion_valida(self):
        configuracion, actualizada = actualizar_configuracion_desde_hoja(
            self.directorio, self.hoja(CONFIG)
        )
        self.assertTrue(actualizada)
        self.assertEqual(configuracion, CONFIG)
        self.assertEqual(cargar_configuracion_local(self.directorio), CONFIG)

    def test_siguiente_inicio_usa_cache_sin_hoja(self):
        self.assertTrue(guardar_configuracion_local(self.directorio, CONFIG))
        self.assertEqual(cargar_configuracion_local(self.directorio), CONFIG)

    def test_configuracion_vacia_o_error_no_destruye_cache(self):
        guardar_configuracion_local(self.directorio, CONFIG)
        configuracion, actualizada = actualizar_configuracion_desde_hoja(
            self.directorio, HojaFalsa([["CLAVE", "VALOR"], ["API_KEY", ""]])
        )
        self.assertFalse(actualizada)
        self.assertEqual(configuracion, CONFIG)
        self.assertEqual(cargar_configuracion_local(self.directorio), CONFIG)

    def test_version_actualizable_reemplaza_cache_completo(self):
        guardar_configuracion_local(self.directorio, CONFIG)
        nueva = dict(CONFIG, VERSION_CONFIG="2")
        configuracion, actualizada = actualizar_configuracion_desde_hoja(
            self.directorio, self.hoja(nueva)
        )
        self.assertTrue(actualizada)
        self.assertEqual(configuracion["VERSION_CONFIG"], "2")

    def test_rechaza_url_que_no_es_endpoint_exec_sin_query_ni_fragmento(self):
        for url in (
            "https://script.google.com/macros/s/ejemplo/dev",
            "https://script.google.com/macros/s/ejemplo/exec?token=secreto",
            "https://script.google.com/macros/s/ejemplo/exec#fragmento",
            "https://otro.example/macros/s/ejemplo/exec",
        ):
            configuracion = dict(CONFIG, URL_APPS_SCRIPT=url)
            self.assertFalse(guardar_configuracion_local(self.directorio, configuracion))

    def test_archivo_local_esta_ignorado_por_git(self):
        raiz = Path(__file__).parent
        resultado = subprocess.run(
            ["git", "check-ignore", "-q", "sincronizacion_config.json"],
            cwd=raiz, check=False,
        )
        self.assertEqual(resultado.returncode, 0)

    def test_archivo_no_contiene_clave_en_otro_lugar(self):
        guardar_configuracion_local(self.directorio, CONFIG)
        contenido = json.loads((self.directorio / "sincronizacion_config.json").read_text("utf-8"))
        self.assertEqual(contenido, CONFIG)


if __name__ == "__main__":
    unittest.main()
