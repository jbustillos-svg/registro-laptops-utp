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


if __name__ == "__main__":
    unittest.main()
