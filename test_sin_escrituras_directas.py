"""Barreras estáticas: Python no puede modificar las hojas operativas."""
import ast
from pathlib import Path
import unittest


class PruebasSinEscriturasDirectas(unittest.TestCase):
    def test_modulos_produccion_no_invocan_escrituras_gspread(self):
        raiz = Path(__file__).parent
        prohibidas = {"update_cell", "batch_update", "append_row", "insert_row"}
        infracciones = []
        for ruta in raiz.glob("*.py*"):
            if ruta.name.startswith("test_"):
                continue
            arbol = ast.parse(ruta.read_text(encoding="utf-8-sig"), filename=str(ruta))
            for nodo in ast.walk(arbol):
                if isinstance(nodo, ast.Call) and isinstance(nodo.func, ast.Attribute):
                    if nodo.func.attr in prohibidas:
                        infracciones.append(f"{ruta.name}:{nodo.lineno}:{nodo.func.attr}")
        self.assertEqual(infracciones, [])

    def test_operaciones_criticas_pasan_por_cliente_unico(self):
        fuente = (Path(__file__).parent / "registro_laptop.pyw").read_text(encoding="utf-8-sig")
        self.assertIn("enviar_evento_coordinador", fuente)
        self.assertIn("registrar_prestamo_coordinador", fuente)


if __name__ == "__main__":
    unittest.main()
