"""Contrato concurrente del coordinador; usa un fake, nunca Google real."""
import threading
import unittest


class CoordinadorFalso:
    def __init__(self, clave):
        self.clave = clave
        self.lock = threading.Lock()
        self.aplicados = {}
        self.no_entregas = 0

    def post(self, evento):
        if evento.get("api_key") != self.clave:
            return {"estado": "ERROR", "codigo": "NO_AUTORIZADO"}
        with self.lock:
            previo = self.aplicados.get(evento["id_evento"])
            if previo:
                return {"estado": "YA_APLICADO", "referencia_registro": previo}
            if evento["tipo_evento"] == "NO_ENTREGA":
                self.no_entregas += 1
            referencia = "Registros!2"
            self.aplicados[evento["id_evento"]] = referencia
            return {"estado": "APLICADO", "referencia_registro": referencia}


class PruebasContratoAppsScript(unittest.TestCase):
    def setUp(self):
        self.coordinador = CoordinadorFalso("clave")

    def test_clave_incorrecta_se_rechaza(self):
        respuesta = self.coordinador.post({"api_key": "mala", "id_evento": "e", "tipo_evento": "PRESTAMO"})
        self.assertEqual(respuesta["codigo"], "NO_AUTORIZADO")

    def test_mismo_uuid_concurrente_se_aplica_una_vez(self):
        evento = {"api_key": "clave", "id_evento": "uuid-1", "tipo_evento": "PRESTAMO"}
        respuestas = []
        hilos = [threading.Thread(target=lambda: respuestas.append(self.coordinador.post(evento))) for _ in range(12)]
        for hilo in hilos: hilo.start()
        for hilo in hilos: hilo.join()
        self.assertEqual(sum(r["estado"] == "APLICADO" for r in respuestas), 1)
        self.assertEqual(sum(r["estado"] == "YA_APLICADO" for r in respuestas), 11)

    def test_no_entrega_concurrente_incrementa_una_vez(self):
        evento = {"api_key": "clave", "id_evento": "uuid-no-entrega", "tipo_evento": "NO_ENTREGA"}
        hilos = [threading.Thread(target=lambda: self.coordinador.post(evento)) for _ in range(12)]
        for hilo in hilos: hilo.start()
        for hilo in hilos: hilo.join()
        self.assertEqual(self.coordinador.no_entregas, 1)


if __name__ == "__main__":
    unittest.main()
