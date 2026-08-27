import ast
from pathlib import Path
import unittest
from unittest import mock

import proximidad_carrito as proximidad
from identidad_roles import rol_requiere_curp


RUTA_APP = Path(__file__).with_name("registro_laptop.pyw")


class PruebasActualizacionProximidad(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fuente = RUTA_APP.read_text(encoding="utf-8-sig")
        cls.arbol = ast.parse(cls.fuente)

    def asignacion(self, nombre):
        for nodo in self.arbol.body:
            if isinstance(nodo, ast.Assign):
                if any(isinstance(destino, ast.Name) and destino.id == nombre
                       for destino in nodo.targets):
                    return ast.literal_eval(nodo.value)
        self.fail(f"No se encontró {nombre}")

    def test_proximidad_esta_habilitada(self):
        self.assertIs(self.asignacion("PROXIMIDAD_HABILITADA"), True)
        self.assertIs(self.asignacion("MODO_PRUEBA"), False)
        self.assertEqual(self.asignacion("VERSION_SISTEMA"), "1.3.8")

    def test_prestamo_confirma_nombre_luego_proximidad_y_luego_identidad(self):
        confirmar = self.fuente.index("confirmado = mostrar_confirmacion_personalizada")
        proximidad = self.fuente.index("threading.Thread(target=comprobar_proximidad", confirmar)
        continuar = self.fuente.index("mostrar_confirmacion_simple(", proximidad)
        self.assertLess(confirmar, proximidad)
        self.assertLess(proximidad, continuar)
        self.assertTrue(rol_requiere_curp("ALUMNO"))
        self.assertFalse(rol_requiere_curp("DOCENTE"))

    def test_devolucion_proximidad_antes_de_sheets_registro_y_apagado(self):
        inicio = self.fuente.index("def entregar_y_apagar")
        proximidad = self.fuente.index("def comprobar_proximidad", inicio)
        sheets = self.fuente.index("def comprobar_devolucion", proximidad)
        localizar = self.fuente.index("fila_encontrada = None", sheets)
        registro = self.fuente.index("mostrar_ventana_espera_registro", localizar)
        apagado = self.fuente.index("apagar_windows()")
        self.assertLess(proximidad, sheets)
        self.assertLess(sheets, localizar)
        self.assertLess(localizar, registro)
        self.assertIn("registrar_salida_con_reintentos", self.fuente)
        self.assertIn("No se confirmó Hora_Salida", self.fuente)
        self.assertGreater(apagado, 0)

    def test_cerca_continua_en_el_nuevo_umbral(self):
        for lectura in (-60, -66, -67):
            with self.subTest(lectura=lectura), \
                    mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=object()), \
                    mock.patch.object(proximidad, "_obtener_rssi", return_value=lectura):
                self.assertEqual(
                    proximidad.verificar_proximidad_carrito(),
                    (True, proximidad.ESTADO_CERCA),
                )

    def test_zona_dudosa_conserva_relecturas_y_confirmacion(self):
        casos = (
            ((-68, -69, -70), (True, proximidad.ESTADO_CERCA)),
            ((-70, -70, -70), (True, proximidad.ESTADO_CERCA)),
            ((-70, -71, -72), (False, proximidad.ESTADO_LEJOS)),
        )
        for lecturas, esperado in casos:
            with self.subTest(lecturas=lecturas), \
                    mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=object()), \
                    mock.patch.object(proximidad, "_obtener_rssi", side_effect=lecturas) as obtener:
                self.assertEqual(proximidad.verificar_proximidad_carrito(), esperado)
                self.assertEqual(obtener.call_count, proximidad.MAX_INTENTOS)

    def test_menos_73_es_lejos_inmediato(self):
        for lectura in (-73, -75):
            with self.subTest(lectura=lectura), \
                    mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=object()), \
                    mock.patch.object(proximidad, "_obtener_rssi", return_value=lectura):
                self.assertEqual(
                    proximidad.verificar_proximidad_carrito(),
                    (False, proximidad.ESTADO_LEJOS),
                )

    def test_no_detectado_y_error_conservan_estados_de_contingencia(self):
        with mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=object()), \
                mock.patch.object(proximidad, "_obtener_rssi", return_value=None):
            self.assertEqual(
                proximidad.verificar_proximidad_carrito(),
                (False, proximidad.ESTADO_NO_DETECTADO),
            )
        with mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=None):
            self.assertEqual(
                proximidad.verificar_proximidad_carrito(),
                (False, proximidad.ESTADO_ERROR),
            )
        bloque = self.fuente[self.fuente.index("def procesar_proximidad"):]
        self.assertIn("ESTADO_NO_DETECTADO", bloque)
        self.assertIn("ESTADO_ERROR", bloque)
        self.assertIn("iniciar_contexto_contingencia", bloque)

    def test_ssid_bssid_y_nuevos_umbrales(self):
        self.assertEqual(proximidad.SSID_CARRITO, "UARB-SYS-01")
        self.assertEqual(proximidad.BSSID_CARRITO, "b8:27:eb:76:e2:f5")
        self.assertEqual(proximidad.UMBRAL_CERCA_SEGURO, -67)
        self.assertEqual(proximidad.UMBRAL_LEJOS_SEGURO, -73)
        self.assertEqual(proximidad.UMBRAL_CONFIRMACION, -70)


if __name__ == "__main__":
    unittest.main()
