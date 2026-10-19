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

    def test_cierres_y_devolucion_dejan_trazas_con_id_sesion(self):
        for evento in (
            "SESION_LOCAL_CIERRE_SOLICITADO ",
            "SESION_LOCAL_CERRADA id_sesion=",
            "SESION_LOCAL_CIERRE_IGNORADO ",
            "SESION_ACTIVA_CONSULTADA id_sesion=",
        ):
            self.assertIn(evento, self.fuente)
        self.assertIn("tipo_evento=DEVOLUCION", self.fuente)
        self.assertIn('id_sesion=sesion_previa["id_sesion"]', self.fuente)

    def test_eventos_durables_solicitan_sincronizacion_inmediata_sin_post_en_ui(self):
        self.assertIn("def solicitar_sincronizacion_inmediata(tipo_evento):", self.fuente)
        for traza in (
            "SINCRONIZACION_INMEDIATA_SOLICITADA tipo=",
            "SINCRONIZACION_INMEDIATA_INICIADA",
            "SINCRONIZACION_INMEDIATA_OMITIDA motivo=ya_en_ejecucion",
        ):
            self.assertIn(traza, self.fuente)

        solicitud = self.fuente.split(
            "def solicitar_sincronizacion_inmediata(tipo_evento):", 1
        )[1].split("def iniciar_sincronizacion_pendientes", 1)[0]
        self.assertIn("sincronizacion_lock.acquire(blocking=False)", solicitud)
        self.assertIn("threading.Thread(target=sincronizar, daemon=True).start()", solicitud)
        self.assertLess(
            solicitud.index("sincronizacion_lock.acquire(blocking=False)"),
            solicitud.index("threading.Thread(target=sincronizar, daemon=True).start()"),
        )

        envio = self.fuente.split("def enviar_evento_local_inmediato(id_evento):", 1)[1].split(
            "def crear_y_enviar_evento_critico", 1
        )[0]
        self.assertIn('solicitar_sincronizacion_inmediata(evento["tipo_evento"])', envio)
        self.assertNotIn("enviar_evento_coordinador(", envio)

        # PRESTAMO y DEVOLUCION despiertan la cola después de persistir SQLite.
        self.assertIn('sesion.get("prestamo_evento_id")', self.fuente)
        self.assertIn('sesion.get("devolucion_evento_id")', self.fuente)
        # NO_ENTREGA y CIERRE_AUTOMATICO comparten el mismo despachador de cola.
        self.assertIn('for campo in ("no_entrega_evento_id", "cierre_automatico_evento_id"):', self.fuente)

    def test_devolucion_durable_solicita_cola_sin_esperar_post(self):
        salida = self.fuente.split("def registrar_salida_con_reintentos(", 1)[1].split(
            "def mostrar_ventana_espera_registro", 1
        )[0]
        self.assertIn("registrar_devolucion_offline(", salida)
        self.assertIn("enviar_evento_local_inmediato(id_evento)", salida)
        self.assertLess(
            salida.index("registrar_devolucion_offline("),
            salida.index("enviar_evento_local_inmediato(id_evento)"),
        )
        self.assertNotIn("enviar_evento_coordinador(", salida)

    def test_sincronizacion_periodica_e_inicial_se_conservan_como_respaldo(self):
        self.assertIn("ventana.after(15000, sincronizar_pendientes_periodicamente)", self.fuente)
        self.assertIn("ventana.after(45000, sincronizar_pendientes_periodicamente)", self.fuente)

    def test_devolucion_pre_apagado_usa_un_intento_breve_del_sincronizador_actual(self):
        self.assertIn("TIMEOUT_SINCRONIZACION_PRE_APAGADO = (2, 3)", self.fuente)
        previa = self.fuente.split("def sincronizar_devolucion_pre_apagado(id_evento):", 1)[1].split(
            "def iniciar_sincronizacion_pendientes", 1
        )[0]
        self.assertIn("sincronizacion_lock.acquire(blocking=False)", previa)
        self.assertIn(".sincronizar_evento(id_evento, timeout=TIMEOUT_SINCRONIZACION_PRE_APAGADO)", previa)
        self.assertIn("DEVOLUCION_PRE_APAGADO_SINCRONIZACION_APLICADA", previa)
        self.assertIn("DEVOLUCION_PRE_APAGADO_SINCRONIZACION_PENDIENTE", previa)

    def test_entrega_guarda_antes_de_sincronizar_y_apagar_sin_doble_worker(self):
        espera = self.fuente.split("def mostrar_ventana_espera_registro(", 1)[1].split(
            "def entregar_y_apagar", 1
        )[0]
        self.assertIn("solicitar_sincronizacion=False", espera)
        self.assertIn('"Registrando entrega..."', espera)
        self.assertIn("sincronizar_devolucion_pre_apagado(id_evento)", espera)
        self.assertNotIn("time.sleep(", espera)
        self.assertLess(
            espera.index("registrar_salida_con_reintentos("),
            espera.index("sincronizar_devolucion_pre_apagado(id_evento)"),
        )
        self.assertLess(
            espera.index("sincronizar_devolucion_pre_apagado(id_evento)"),
            espera.index("apagar_windows()"),
        )
        self.assertIn("if exito:", espera)
        self.assertIn("if not sesion:\n            return False", espera)
        self.assertIn("state=\"disabled\"", self.fuente)


if __name__ == "__main__":
    unittest.main()
