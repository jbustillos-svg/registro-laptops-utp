import time
from statistics import median

import pywifi


SSID_CARRITO = "UARB-SYS-01"
BSSID_CARRITO = "b8:27:eb:76:e2:f5"

UMBRAL_CERCA_SEGURO = -69
UMBRAL_LEJOS_SEGURO = -75
UMBRAL_CONFIRMACION = -72

ESPERA_RESULTADOS_SCAN = 0.80
MAX_INTENTOS = 5

ESTADO_CERCA = "CERCA"
ESTADO_LEJOS = "LEJOS"
ESTADO_NO_DETECTADO = "NO_DETECTADO"
ESTADO_ERROR = "ERROR"


def _normalizar_bssid(bssid):
    return (bssid or "").strip().lower().rstrip(":")


def _obtener_interfaz_wifi():
    wifi = pywifi.PyWiFi()
    interfaces = wifi.interfaces()

    if not interfaces:
        return None

    return interfaces[0]


def _extraer_rssi(resultados):
    lecturas = []

    for red in resultados:
        ssid = (red.ssid or "").strip()
        bssid = _normalizar_bssid(red.bssid)

        if ssid == SSID_CARRITO and bssid == BSSID_CARRITO:
            lecturas.append(red.signal)

    if not lecturas:
        return None

    return max(lecturas)


def _obtener_rssi(iface):
    inicio_scan = time.perf_counter()
    iface.scan()
    print("[SCAN] Solicitud enviada")

    time.sleep(ESPERA_RESULTADOS_SCAN)
    lectura = _extraer_rssi(iface.scan_results())
    print(f"[SCAN] +{time.perf_counter() - inicio_scan:.2f} s -> RSSI {lectura}")
    return lectura


def verificar_proximidad_carrito(
    mostrar_detalles=False,
    callback_intento=None,
    callback_evento=None,
):
    inicio = time.perf_counter()

    def finalizar(resultado, estado, descripcion):
        print(f"[PROXIMIDAD] Resultado: {descripcion}")
        print(f"[PROXIMIDAD] Tiempo total: {time.perf_counter() - inicio:.2f} s")
        return resultado, estado

    def registrar(evento):
        print(f"[PROXIMIDAD] {evento}")
        if callback_evento:
            try:
                callback_evento(f"PROXIMIDAD {evento}")
            except Exception as error_callback:
                print(f"[PROXIMIDAD] Error registrando evento: {error_callback}")

    try:
        iface = _obtener_interfaz_wifi()
        if iface is None:
            return finalizar(False, ESTADO_ERROR, "SIN ADAPTADOR WI-FI")

        lecturas = []
        intentos_con_error = 0

        for intento in range(1, MAX_INTENTOS + 1):
            print(f"[PROXIMIDAD] Intento {intento}/{MAX_INTENTOS}")
            if callback_intento:
                try:
                    callback_intento(intento, MAX_INTENTOS)
                except Exception as error_callback:
                    print(f"[PROXIMIDAD] Error actualizando progreso: {error_callback}")

            try:
                lectura = _obtener_rssi(iface)
            except Exception as error:
                intentos_con_error += 1
                registrar(
                    f"intento={intento}/{MAX_INTENTOS} escaneo=nuevo "
                    f"resultado=ERROR detalle={type(error).__name__}"
                )
                if mostrar_detalles:
                    import traceback
                    traceback.print_exc()
                continue

            if lectura is None:
                registrar(
                    f"intento={intento}/{MAX_INTENTOS} escaneo=nuevo "
                    "resultado=NO_DETECTADO"
                )
                continue

            lecturas.append(lectura)
            registrar(
                f"intento={intento}/{MAX_INTENTOS} escaneo=nuevo "
                f"BSSID={BSSID_CARRITO} RSSI={lectura}"
            )

            if lectura >= UMBRAL_CERCA_SEGURO:
                return finalizar(True, ESTADO_CERCA, "CERCA CLARAMENTE")

            if lectura <= UMBRAL_LEJOS_SEGURO:
                return finalizar(False, ESTADO_LEJOS, "LEJOS CLARAMENTE")

            print("[PROXIMIDAD] Zona dudosa")

        if not lecturas:
            if intentos_con_error == MAX_INTENTOS:
                return finalizar(False, ESTADO_ERROR, "ERROR EN TODOS LOS INTENTOS")
            return finalizar(
                False,
                ESTADO_NO_DETECTADO,
                "SIN SSID/BSSID VÁLIDO"
            )

        mediana = median(lecturas)
        print(f"[PROXIMIDAD] Mediana: {mediana:.1f} dBm")
        autorizado = mediana >= UMBRAL_CONFIRMACION
        estado = ESTADO_CERCA if autorizado else ESTADO_LEJOS
        return finalizar(autorizado, estado, estado)

    except Exception as error:
        if mostrar_detalles:
            print(f"[PROXIMIDAD] Error de lectura: {error}")
        return finalizar(False, ESTADO_ERROR, "ERROR DE LECTURA")


if __name__ == "__main__":
    print("\nVerificando proximidad del carrito...\n")
    autorizado, estado = verificar_proximidad_carrito(mostrar_detalles=True)
    print("\n------------------------------")

    if autorizado:
        print("AUTORIZADO")
        print("El equipo se encuentra dentro del área permitida del carrito.")
    else:
        print("BLOQUEADO")
        print("Acérquese al carrito móvil para continuar.")

    print(f"Estado: {estado}")

    print("------------------------------")
