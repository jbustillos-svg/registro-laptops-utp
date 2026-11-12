import ast
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest import mock

import proximidad_carrito as proximidad
from identidad_roles import rol_requiere_curp


RUTA_APP = Path(__file__).with_name("registro_laptop.pyw")


def red_carrito(rssi):
    return SimpleNamespace(
        ssid=proximidad.SSID_CARRITO,
        bssid=proximidad.BSSID_CARRITO,
        signal=rssi,
    )


class InterfazSimulada:
    def __init__(self, resultados_por_scan):
        self.resultados_por_scan = iter(resultados_por_scan)
        self.resultado_actual = []
        self.llamadas_scan = 0
        self.llamadas_resultados = 0

    def scan(self):
        self.llamadas_scan += 1
        resultado = next(self.resultados_por_scan)
        if isinstance(resultado, Exception):
            raise resultado
        self.resultado_actual = resultado

    def scan_results(self):
        self.llamadas_resultados += 1
        return self.resultado_actual


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

    def test_proximidad_esta_deshabilitada_temporalmente(self):
        self.assertIs(self.asignacion("PROXIMIDAD_HABILITADA"), False)
        self.assertIs(self.asignacion("MODO_PRUEBA"), False)
        self.assertEqual(self.asignacion("VERSION_SISTEMA"), "1.4.12")

    def test_prestamo_omite_escaneo_y_continua_a_identidad_si_esta_deshabilitada(self):
        confirmar = self.fuente.index("confirmado = mostrar_confirmacion_personalizada")
        omision = self.fuente.index(
            "PROXIMIDAD_OMITIDA motivo=DESHABILITADA operacion=PRESTAMO", confirmar
        )
        continuar = self.fuente.index("mostrar_confirmacion_simple(", omision)
        escaneo = self.fuente.index("threading.Thread(target=comprobar_proximidad", continuar)
        self.assertLess(confirmar, omision)
        self.assertLess(omision, continuar)
        self.assertLess(continuar, escaneo)
        self.assertTrue(rol_requiere_curp("ALUMNO"))
        self.assertFalse(rol_requiere_curp("DOCENTE"))

    def test_devolucion_omite_proximidad_y_entra_a_pipeline_durable(self):
        bloque = self.fuente.split("    def iniciar_devolucion():", 1)[1].split(
            "    ventana._abortar_devolucion", 1
        )[0]
        omision = bloque.index(
            "PROXIMIDAD_OMITIDA motivo=DESHABILITADA operacion=DEVOLUCION"
        )
        durable = bloque.index("iniciar_persistencia()", omision)
        modal = bloque.index("crear_modal_verificacion_ubicacion(", durable)
        ruta_habilitada = bloque.index("threading.Thread(target=worker_proximidad", modal)
        self.assertLess(omision, durable)
        self.assertLess(durable, modal)
        self.assertLess(modal, ruta_habilitada)

    def test_bandera_true_conserva_ruta_de_proximidad(self):
        fuente_habilitada = self.fuente.replace(
            "PROXIMIDAD_HABILITADA = False", "PROXIMIDAD_HABILITADA = True", 1
        )
        self.assertIn("PROXIMIDAD_HABILITADA = True", fuente_habilitada)
        bloque = self.fuente.split("    def iniciar_devolucion():", 1)[1].split(
            "    ventana._abortar_devolucion", 1
        )[0]
        self.assertIn("if not PROXIMIDAD_HABILITADA:", bloque)
        self.assertIn("threading.Thread(target=worker_proximidad, daemon=True).start()", bloque)
        self.assertIn(
            "verificar_proximidad_carrito(callback_evento=registrar_evento_tecnico)",
            self.fuente,
        )

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
        for lectura in (-60, -68, -69):
            with self.subTest(lectura=lectura), \
                    mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=object()), \
                    mock.patch.object(proximidad, "_obtener_rssi", return_value=lectura):
                self.assertEqual(
                    proximidad.verificar_proximidad_carrito(),
                    (True, proximidad.ESTADO_CERCA),
                )

    def test_zona_dudosa_conserva_relecturas_y_confirmacion(self):
        casos = (
            ((-70, -71, -72, -73, -74), (True, proximidad.ESTADO_CERCA)),
            ((-72, -72, -72, -72, -72), (True, proximidad.ESTADO_CERCA)),
            ((-71, -72, -73, -74, -74), (False, proximidad.ESTADO_LEJOS)),
        )
        for lecturas, esperado in casos:
            with self.subTest(lecturas=lecturas), \
                    mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=object()), \
                    mock.patch.object(proximidad, "_obtener_rssi", side_effect=lecturas) as obtener:
                self.assertEqual(proximidad.verificar_proximidad_carrito(), esperado)
                self.assertEqual(obtener.call_count, proximidad.MAX_INTENTOS)

    def test_no_detectado_temporal_continua_con_escaneos_nuevos(self):
        casos = (
            ([[], [red_carrito(-68)]], 2),
            ([[], [], [], [], [red_carrito(-68)]], 5),
        )
        for resultados, escaneos_esperados in casos:
            interfaz = InterfazSimulada(resultados)
            intentos = []
            eventos = []
            with self.subTest(escaneos=escaneos_esperados), \
                    mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=interfaz), \
                    mock.patch.object(proximidad.time, "sleep") as espera:
                self.assertEqual(
                    proximidad.verificar_proximidad_carrito(
                        callback_intento=lambda intento, total: intentos.append((intento, total)),
                        callback_evento=eventos.append,
                    ),
                    (True, proximidad.ESTADO_CERCA),
                )
            self.assertEqual(interfaz.llamadas_scan, escaneos_esperados)
            self.assertEqual(interfaz.llamadas_resultados, escaneos_esperados)
            self.assertEqual(espera.call_count, escaneos_esperados)
            self.assertEqual(intentos[-1], (escaneos_esperados, 5))
            self.assertIn("resultado=NO_DETECTADO", eventos[0])

    def test_cinco_intentos_vacios_devuelven_no_detectado(self):
        interfaz = InterfazSimulada([[] for _ in range(5)])
        with mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=interfaz), \
                mock.patch.object(proximidad.time, "sleep"):
            self.assertEqual(
                proximidad.verificar_proximidad_carrito(),
                (False, proximidad.ESTADO_NO_DETECTADO),
            )
        self.assertEqual(interfaz.llamadas_scan, 5)
        self.assertEqual(interfaz.llamadas_resultados, 5)

    def test_error_temporal_continua_con_el_siguiente_escaneo(self):
        interfaz = InterfazSimulada([RuntimeError("temporal"), [red_carrito(-68)]])
        eventos = []
        with mock.patch.object(proximidad, "_obtener_interfaz_wifi", return_value=interfaz), \
                mock.patch.object(proximidad.time, "sleep"):
            self.assertEqual(
                proximidad.verificar_proximidad_carrito(callback_evento=eventos.append),
                (True, proximidad.ESTADO_CERCA),
            )
        self.assertEqual(interfaz.llamadas_scan, 2)
        self.assertEqual(interfaz.llamadas_resultados, 1)
        self.assertIn("intento=1/5 escaneo=nuevo resultado=ERROR", eventos[0])

    def test_cada_intento_es_scan_espera_y_una_lista_nueva(self):
        interfaz = InterfazSimulada([[], [], [], [], []])
        with mock.patch.object(proximidad.time, "sleep") as espera:
            for _ in range(5):
                self.assertIsNone(proximidad._obtener_rssi(interfaz))
        self.assertEqual(interfaz.llamadas_scan, 5)
        self.assertEqual(interfaz.llamadas_resultados, 5)
        espera.assert_has_calls(
            [mock.call(proximidad.ESPERA_RESULTADOS_SCAN) for _ in range(5)]
        )

    def test_contador_visual_y_workers_conservan_interfaz_responsiva(self):
        self.assertEqual(proximidad.MAX_INTENTOS, 5)
        self.assertIn('estado="Intento 1 de 5"', self.fuente)
        self.assertNotIn("Intentando 1 de 3", self.fuente)
        self.assertEqual(
            self.fuente.count('text=f"Intento {intento} de {total}"'),
            2,
        )
        self.assertEqual(
            self.fuente.count("threading.Thread(target=comprobar_proximidad, daemon=True).start()"),
            2,
        )
        self.assertEqual(
            self.fuente.count("callback_evento=registrar_evento_tecnico"),
            2,
        )

    def test_menos_75_es_lejos_inmediato(self):
        for lectura in (-75, -76, -80):
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
        self.assertEqual(proximidad.UMBRAL_CERCA_SEGURO, -69)
        self.assertEqual(proximidad.UMBRAL_LEJOS_SEGURO, -75)
        self.assertEqual(proximidad.UMBRAL_CONFIRMACION, -72)


if __name__ == "__main__":
    unittest.main()
