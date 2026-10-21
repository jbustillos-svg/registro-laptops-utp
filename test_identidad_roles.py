import ast
from pathlib import Path
import unittest

from identidad_roles import (
    normalizar_rol,
    rol_requiere_curp,
    seleccionar_flujo_identidad,
)


class ErrorConsultaGoogle(Exception):
    pass


def cargar_consultar_datos_identidad(hoja):
    ruta = Path(__file__).with_name("registro_laptop.pyw")
    arbol = ast.parse(ruta.read_text(encoding="utf-8-sig"))
    funcion = next(
        nodo for nodo in arbol.body
        if isinstance(nodo, ast.FunctionDef)
        and nodo.name == "consultar_datos_identidad"
    )
    modulo = ast.Module(body=[funcion], type_ignores=[])
    ast.fix_missing_locations(modulo)
    espacio = {
        "ErrorConsultaGoogle": ErrorConsultaGoogle,
        "hoja_alumnos": hoja,
        "normalizar_rol": normalizar_rol,
        "seleccionar_flujo_identidad": seleccionar_flujo_identidad,
        "rol_requiere_curp": rol_requiere_curp,
    }
    exec(compile(modulo, str(ruta), "exec"), espacio)
    return espacio["consultar_datos_identidad"]


class Celda:
    def __init__(self, valor):
        self.value = valor


class HojaFalsa:
    def __init__(self, filas, error=False):
        self.filas = filas
        self.error = error
        self.columnas_leidas = []

    def col_values(self, columna):
        if self.error:
            raise RuntimeError("Sheets no disponible")
        return [fila[columna - 1] for fila in self.filas]

    def cell(self, fila, columna):
        self.columnas_leidas.append(columna)
        return Celda(self.filas[fila - 1][columna - 1])


class PruebasRolesIdentidad(unittest.TestCase):
    def test_docente_vacio_omite_curp(self):
        hoja = HojaFalsa([["082025", "Docentes y Administrativos", "DOCENTE", ""]])
        consultar = cargar_consultar_datos_identidad(hoja)
        self.assertEqual(consultar("082025"), ("DOCENTE", None))
        self.assertNotIn(4, hoja.columnas_leidas)
        self.assertFalse(rol_requiere_curp("DOCENTE"))

    def test_otro_docente_vacio_tambien_omite_curp(self):
        hoja = HojaFalsa([["D-2", "Otro docente", " docente ", ""]])
        consultar = cargar_consultar_datos_identidad(hoja)
        self.assertEqual(consultar("D-2"), ("DOCENTE", None))

    def test_docente_con_curp_no_es_obligado_a_usarla(self):
        hoja = HojaFalsa([["D-3", "Docente", "DOCENTE", "CURP1234"]])
        consultar = cargar_consultar_datos_identidad(hoja)
        self.assertEqual(consultar("D-3"), ("DOCENTE", None))
        self.assertNotIn(4, hoja.columnas_leidas)

    def test_alumno_con_curp_la_consulta(self):
        hoja = HojaFalsa([["A-1", "Alumno", "ALUMNO", "CURP5678"]])
        consultar = cargar_consultar_datos_identidad(hoja)
        self.assertEqual(consultar("A-1"), ("ALUMNO", "CURP5678"))
        self.assertIn(4, hoja.columnas_leidas)
        self.assertTrue(rol_requiere_curp("ALUMNO"))

    def test_alumno_vacio_no_cambia_de_flujo(self):
        hoja = HojaFalsa([["A-2", "Alumno", "", ""]])
        consultar = cargar_consultar_datos_identidad(hoja)
        self.assertEqual(consultar("A-2"), ("ALUMNO", ""))
        self.assertEqual(seleccionar_flujo_identidad(""), "ALUMNO")

    def test_error_real_de_sheets_se_propaga(self):
        consultar = cargar_consultar_datos_identidad(HojaFalsa([], error=True))
        with self.assertRaises(ErrorConsultaGoogle):
            consultar("A-1")

    def test_matricula_inexistente_permanece_no_encontrada(self):
        consultar = cargar_consultar_datos_identidad(HojaFalsa([]))
        self.assertEqual(consultar("NO-EXISTE"), (None, None))

    def test_solo_docente_usa_flujo_docente(self):
        self.assertEqual(normalizar_rol(" docente "), "DOCENTE")
        self.assertEqual(seleccionar_flujo_identidad("ADMINISTRATIVO"), "ALUMNO")


if __name__ == "__main__":
    unittest.main()
