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

    def test_prestamo_reconcilia_y_normaliza_antes_de_insertar(self):
        for simbolo in (
            "buscarRegistrosPrestamoCompatibles_",
            "normalizarFecha_",
            "normalizarHora_",
            "DUPLICADOS_PREEXISTENTES",
            "VERIFICACION_PRESTAMO_FALLIDA",
        ):
            self.assertIn(simbolo, self.fuente)

    def test_matricula_es_texto_y_devolucion_verifica_horas_normalizadas(self):
        for simbolo in (
            "setNumberFormat('@')",
            "setHorizontalAlignment('left')",
            "matriculaEquivalentePorCeroInicial_",
            "buscarRegistroPorSesion_",
            "horaSalidaCoincide_",
            "MATRICULA_CERO_INICIAL_RECUPERADA",
            "Math.floor(valor) * 86400000",
        ):
            self.assertIn(simbolo, self.fuente)
        self.assertNotIn("parseInt(evento.matricula", self.fuente)
        self.assertNotIn("Number(evento.matricula", self.fuente)


if __name__ == "__main__":
    unittest.main()
