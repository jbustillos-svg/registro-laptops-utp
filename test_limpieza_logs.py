import ast
import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest import mock


def cargar_limpiador(directorio, registrar=lambda evento: None):
    ruta = Path(__file__).with_name("registro_laptop.pyw")
    arbol = ast.parse(ruta.read_text(encoding="utf-8-sig"))
    funcion = next(
        nodo for nodo in arbol.body
        if isinstance(nodo, ast.FunctionDef)
        and nodo.name == "limpiar_logs_antiguos"
    )
    modulo = ast.Module(body=[funcion], type_ignores=[])
    ast.fix_missing_locations(modulo)
    espacio = {
        "DIRECTORIO_APP": str(directorio),
        "os": os,
        "time": time,
        "registrar_evento_tecnico": registrar,
    }
    exec(compile(modulo, str(ruta), "exec"), espacio)
    return espacio["limpiar_logs_antiguos"]


class PruebasLimpiezaLogs(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.directorio = Path(self.temporal.name)
        self.ahora = time.time()

    def tearDown(self):
        self.temporal.cleanup()

    def crear(self, nombre, dias):
        ruta = self.directorio / nombre
        ruta.write_text("prueba", encoding="utf-8")
        marca = self.ahora - dias * 24 * 60 * 60
        os.utime(ruta, (marca, marca))
        return ruta

    def test_retiene_hasta_siete_dias_y_elimina_solo_log_mas_antiguo(self):
        hoy = self.crear("hoy.log", 0)
        seis = self.crear("seis.log", 6)
        siete = self.crear("siete.log", 7)
        antiguo = self.crear("antiguo.log", 8)
        txt = self.crear("antiguo.txt", 30)
        json = self.crear("antiguo.json", 30)
        py = self.crear("antiguo.py", 30)

        limpiar = cargar_limpiador(self.directorio)
        with mock.patch.object(time, "time", return_value=self.ahora):
            self.assertEqual(limpiar(), 1)

        self.assertTrue(hoy.exists())
        self.assertTrue(seis.exists())
        self.assertTrue(siete.exists())
        self.assertFalse(antiguo.exists())
        self.assertTrue(txt.exists())
        self.assertTrue(json.exists())
        self.assertTrue(py.exists())

    def test_error_de_borrado_no_bloquea_y_continua(self):
        bloqueado = self.crear("bloqueado.log", 8)
        eliminable = self.crear("eliminable.log", 9)
        eventos = []
        limpiar = cargar_limpiador(self.directorio, eventos.append)
        eliminar_real = os.remove

        def eliminar(ruta):
            if Path(ruta) == bloqueado:
                raise PermissionError("en uso")
            eliminar_real(ruta)

        with mock.patch.object(time, "time", return_value=self.ahora), \
                mock.patch.object(os, "remove", side_effect=eliminar):
            self.assertEqual(limpiar(), 1)

        self.assertTrue(bloqueado.exists())
        self.assertFalse(eliminable.exists())
        self.assertEqual(eventos, ["LIMPIEZA_LOGS eliminados=1"])

    def test_no_entra_en_subdirectorios(self):
        subdirectorio = self.directorio / ".venv"
        subdirectorio.mkdir()
        interno = subdirectorio / "interno.log"
        interno.write_text("prueba", encoding="utf-8")
        marca = self.ahora - 30 * 24 * 60 * 60
        os.utime(interno, (marca, marca))

        limpiar = cargar_limpiador(self.directorio)
        with mock.patch.object(time, "time", return_value=self.ahora):
            self.assertEqual(limpiar(), 0)
        self.assertTrue(interno.exists())


if __name__ == "__main__":
    unittest.main()
