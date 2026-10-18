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

    def test_arranque_recupera_sesion_local_y_muestra_devolucion(self):
        self.assertIn("def recuperar_sesion_activa_al_iniciar():", self.fuente)
        self.assertIn('sesion = obtener_sesion_local_activa()', self.fuente)
        self.assertIn(
            'lambda: mostrar_ventana_entrega(nombre, matricula, sesion_recuperada=True)',
            self.fuente,
        )
        self.assertIn(
            "ventana.after_idle(recuperar_sesion_activa_al_iniciar)", self.fuente,
        )

    def test_login_no_inicia_otro_prestamo_si_hay_sesion_local_activa(self):
        inicio_sesion = self.fuente.split("def iniciar_sesion():", 1)[1].split(
            "def crear_pantalla_login():", 1
        )[0]
        self.assertIn("if recuperar_sesion_activa_al_iniciar():", inicio_sesion)
        self.assertIn("if recuperar_sesion_activa_al_iniciar():\n        return", inicio_sesion)

    def test_ventana_recuperada_indica_prestamo_activo_y_ofrece_entrega(self):
        self.assertIn("def mostrar_ventana_entrega(nombre, matricula, sesion_recuperada=False):", self.fuente)
        self.assertIn("Esta laptop tiene un préstamo activo.", self.fuente)
        self.assertIn('text="ENTREGAR Y APAGAR"', self.fuente)


if __name__ == "__main__":
    unittest.main()
