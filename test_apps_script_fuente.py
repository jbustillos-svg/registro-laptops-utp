from pathlib import Path
import unittest


class PruebasFuenteAppsScript(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fuente = (Path(__file__).parent / "apps_script_sincronizacion.gs").read_text(
            encoding="utf-8"
        )

    def test_usa_lock_y_liberacion_finalmente(self):
        self.assertIn("LockService.getScriptLock()", self.fuente)
        self.assertIn("finally", self.fuente)
        self.assertIn("releaseLock", self.fuente)

    def test_bitacora_tiene_campos_minimos_y_estados_idempotentes(self):
        for campo in (
            "ID_EVENTO", "TIPO", "MATRICULA", "LAPTOP_ID", "FECHA_HORA_EVENTO",
            "ESTADO", "FECHA_PREPARACION", "FECHA_APLICACION",
            "REFERENCIA_REGISTRO", "DATOS_CONTROL", "EQUIPO_ORIGEN",
            "PREPARANDO", "APLICADO", "YA_APLICADO",
        ):
            self.assertIn(campo, self.fuente)

    def test_no_entrega_usa_batch_atomico(self):
        self.assertIn("Sheets.Spreadsheets.Values.batchUpdate", self.fuente)

    def test_secreto_solo_se_resuelve_desde_script_properties(self):
        self.assertIn("getScriptProperties().getProperty('UTP_SYNC_API_KEY')", self.fuente)
        self.assertNotIn("UTP_SYNC_API_KEY =", self.fuente)

    def test_preparacion_solo_acepta_exec_y_conserva_la_existente(self):
        self.assertIn("esUrlWebAppProduccion_(urlExistente)", self.fuente)
        self.assertIn("esUrlWebAppProduccion_(urlDetectada)", self.fuente)
        self.assertIn("CONFIGURAR_URL_APPS_SCRIPT_EXEC", self.fuente)


if __name__ == "__main__":
    unittest.main()
