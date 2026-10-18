from pathlib import Path
import unittest


class PruebasRegistroLaptopFuente(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fuente = (Path(__file__).parent / "registro_laptop.pyw").read_text(
            encoding="utf-8"
        )

    def test_prestamo_local_registra_todas_las_etapas_tecnicas(self):
        for evento in (
            "PRESTAMO_LOCAL_INICIANDO",
            "PRESTAMO_LOCAL_SESION_EXISTENTE",
            "PRESTAMO_LOCAL_EVENTO_CREADO",
            "PRESTAMO_LOCAL_SESION_CREADA",
            "PRESTAMO_LOCAL_ERROR",
        ):
            self.assertIn(evento, self.fuente)

    def test_error_local_registra_tipo_y_mensaje(self):
        self.assertIn(
            'f"PRESTAMO_LOCAL_ERROR tipo={type(e).__name__} mensaje={str(e)[:300]}"',
            self.fuente,
        )

    def test_arranque_detecta_sesion_local_y_conserva_login_de_matricula(self):
        self.assertIn("def registrar_sesion_pendiente_al_iniciar():", self.fuente)
        self.assertIn('sesion = obtener_sesion_local_activa()', self.fuente)
        self.assertIn(
            "RECUPERACION_SESION_LOCAL resultado=ACTIVA_UI_LOGIN",
            self.fuente,
        )
        self.assertIn(
            "ventana.after_idle(registrar_sesion_pendiente_al_iniciar)", self.fuente,
        )
        recuperacion = self.fuente.split(
            "def registrar_sesion_pendiente_al_iniciar():", 1
        )[1].split("def registrar_devolucion_contingencia", 1)[0]
        self.assertNotIn("ventana.withdraw()", recuperacion)
        self.assertNotIn("mostrar_ventana_entrega", recuperacion)

    def test_misma_matricula_reutiliza_sesion_y_muestra_entrega(self):
        self.assertIn('resultado == "misma_matricula"', self.fuente)
        self.assertIn("RECUPERACION_SESION_LOCAL resultado=MISMA_MATRICULA", self.fuente)
        self.assertIn("lambda: mostrar_ventana_entrega(nombre, matricula)", self.fuente)

    def test_matricula_diferente_resuelve_sesion_por_sqlite_y_apps_script(self):
        self.assertIn("cerrar_sesion_contingencia_por_nueva_sesion(matricula)", self.fuente)
        self.assertIn("resolver_sesion_activa_por_nueva_matricula(", self.fuente)
        self.assertIn("enviar_evento_local_inmediato(id_evento)", self.fuente)
        self.assertIn("registrar_no_entrega=not matricula_exenta_no_entrega", self.fuente)

    def test_devolucion_normal_sigue_disponible(self):
        self.assertIn("def mostrar_ventana_entrega(nombre, matricula):", self.fuente)
        self.assertIn('text="ENTREGAR Y APAGAR"', self.fuente)


if __name__ == "__main__":
    unittest.main()
