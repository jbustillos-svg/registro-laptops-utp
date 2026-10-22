import ast
from pathlib import Path
from types import SimpleNamespace
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

    def test_sesion_residual_se_reconcilia_antes_de_elegir_flujo(self):
        seleccion = self.fuente.split("def mostrar_confirmacion_simple", 1)[1].split(
            "def cambiar_estado", 1
        )[0]
        self.assertIn("SESION_RESIDUAL_DETECTADA", seleccion)
        self.assertIn("reconciliar_sesion_residual_local", seleccion)
        self.assertIn("reconciliar_sesion_residual_remota", seleccion)
        self.assertIn("SESION_RESIDUAL_RECONCILIADA origen=LOCAL", seleccion)
        self.assertIn("SESION_RESIDUAL_RECONCILIADA origen=REMOTO", seleccion)
        self.assertIn("FLUJO_SELECCIONADO tipo=PRESTAMO_NUEVO", seleccion)
        self.assertIn("FLUJO_SELECCIONADO tipo=RECUPERACION_SESION", seleccion)
        self.assertLess(
            seleccion.index("FLUJO_SELECCIONADO tipo=PRESTAMO_NUEVO"),
            seleccion.index("iniciar_validacion_identidad()"),
        )

    def test_devolucion_restauracion_limpia_estado_antes_de_sqlite(self):
        entrega = self.fuente.split("def entregar_y_apagar", 1)[1].split(
            "def mostrar_ventana_entrega", 1
        )[0]
        self.assertIn("def restaurar_estado_devolucion(", entrega)
        self.assertIn('estado_devolucion["en_proceso"] = False', entrega)
        self.assertIn('estado_devolucion["afters"].clear()', entrega)
        self.assertIn('command=estado_devolucion["comando"]', entrega)
        self.assertIn("DEVOLUCION_CLIC_IGNORADO motivo=EN_PROCESO_O_DURABLE", entrega)
        self.assertIn("marcar_durable=marcar_devolucion_durable", entrega)

    def test_matricula_diferente_resuelve_sesion_por_sqlite_y_apps_script(self):
        self.assertIn("cerrar_sesion_contingencia_por_nueva_sesion(matricula)", self.fuente)
        self.assertIn("resolver_sesion_activa_por_nueva_matricula(", self.fuente)
        self.assertIn("enviar_evento_local_inmediato(id_evento)", self.fuente)
        self.assertIn("registrar_no_entrega=not matricula_exenta_no_entrega", self.fuente)

    def test_devolucion_normal_sigue_disponible(self):
        self.assertIn("def mostrar_ventana_entrega(nombre, matricula):", self.fuente)
        self.assertIn('text="ENTREGAR Y APAGAR"', self.fuente)

    def test_identidad_offline_exige_el_verificador_curp_minimo_cacheado(self):
        consulta = self.fuente.split("def consultar_datos_identidad", 1)[1].split(
            "def buscar_curp", 1
        )[0]
        flujo_identidad = self.fuente.split("def mostrar_confirmacion_simple", 1)[1].split(
            "def formatear_fecha", 1
        )[0]
        self.assertIn("curp_ultimos_2 = str(curp).strip()[-2:]", consulta)
        self.assertIn("actualizar_cache_persona(", consulta)
        self.assertIn('persona_cache.get("curp_ultimos_2")', flujo_identidad)
        self.assertIn("Validación offline no disponible", flujo_identidad)
        self.assertNotIn("curp_real = persona_cache.get(\"curp\")", flujo_identidad)

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
        self.assertIn("def solicitar_sincronizacion_inmediata(tipo_evento, id_evento=None):", self.fuente)
        for traza in (
            "SINCRONIZACION_INMEDIATA_SOLICITADA tipo=",
            "SINCRONIZACION_INMEDIATA_INICIADA",
            "SINCRONIZACION_INMEDIATA_OMITIDA motivo=ya_en_ejecucion",
        ):
            self.assertIn(traza, self.fuente)

        solicitud = self.fuente.split(
            "def solicitar_sincronizacion_inmediata(tipo_evento, id_evento=None):", 1
        )[1].split("def iniciar_sincronizacion_pendientes", 1)[0]
        self.assertIn("sincronizacion_lock.acquire(blocking=False)", solicitud)
        self.assertIn("threading.Thread(target=sincronizar, daemon=True).start()", solicitud)
        self.assertIn("sincronizador.sincronizar_evento(id_evento)", solicitud)
        self.assertLess(
            solicitud.index("sincronizacion_lock.acquire(blocking=False)"),
            solicitud.index("threading.Thread(target=sincronizar, daemon=True).start()"),
        )

        envio = self.fuente.split("def enviar_evento_local_inmediato(id_evento):", 1)[1].split(
            "def crear_y_enviar_evento_critico", 1
        )[0]
        self.assertIn('solicitar_sincronizacion_inmediata(evento["tipo_evento"], id_evento=id_evento)', envio)
        self.assertNotIn("enviar_evento_coordinador(", envio)

        # PRESTAMO y DEVOLUCION despiertan la cola después de persistir SQLite.
        self.assertIn('sesion.get("prestamo_evento_id")', self.fuente)
        self.assertIn('sesion.get("devolucion_evento_id")', self.fuente)
        # NO_ENTREGA y CIERRE_AUTOMATICO comparten el mismo despachador de cola.
        self.assertIn('for campo in ("no_entrega_evento_id", "cierre_automatico_evento_id"):', self.fuente)

        inmediata = self.fuente.split("def sincronizar_evento_inmediato", 1)[1].split(
            "def crear_y_enviar_evento_critico", 1
        )[0]
        self.assertIn("SincronizadorContingencia", inmediata)
        self.assertIn(".sincronizar_evento(id_evento, timeout=timeout)", inmediata)
        self.assertIn("ESPERA_LOCK_SINCRONIZACION_INMEDIATA", inmediata)
        self.assertNotIn("enviar_evento_coordinador(", inmediata)

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

    def test_sondeo_de_salud_de_sheets_no_agota_la_cuota(self):
        """La cola inmediata atiende eventos nuevos; este sondeo sólo es respaldo."""
        sondeo = self.fuente.split("def verificar_conexion_periodicamente():", 1)[1].split(
            "def detener_verificacion_conexion", 1
        )[0]
        self.assertIn("time.sleep(60)", sondeo)
        self.assertNotIn("time.sleep(2)", sondeo)

    def test_devolucion_pre_apagado_usa_un_intento_breve_del_sincronizador_actual(self):
        self.assertIn("TIMEOUT_SINCRONIZACION_PRE_APAGADO = (2, 3)", self.fuente)
        previa = self.fuente.split("def sincronizar_devolucion_pre_apagado(id_evento):", 1)[1].split(
            "def iniciar_sincronizacion_pendientes", 1
        )[0]
        self.assertIn("sincronizar_evento_inmediato(", previa)
        self.assertIn("timeout=TIMEOUT_SINCRONIZACION_PRE_APAGADO", previa)
        self.assertIn("DEVOLUCION_PRE_APAGADO_SINCRONIZACION_APLICADA", previa)
        self.assertIn("DEVOLUCION_PRE_APAGADO_SINCRONIZACION_PENDIENTE", previa)

    def test_modo_prueba_simula_apagado_sin_llamar_shutdown(self):
        apagado = self.fuente.split("def apagar_windows():", 1)[1].split(
            "def contar_usos_alumno", 1
        )[0]
        rama_prueba = apagado.split("if MODO_PRUEBA:", 1)[1].split(
            'os.system("shutdown /s /t 3")', 1
        )[0]
        self.assertIn('APAGADO_SIMULADO modo_prueba=True', rama_prueba)
        self.assertIn("crear_pantalla_login()", rama_prueba)
        self.assertNotIn('os.system("shutdown /s /t 3")', rama_prueba)
        self.assertIn('MODO_PRUEBA = False', self.fuente)
        self.assertIn('os.system("shutdown /s /t 3")', apagado)

    def test_entrega_guarda_antes_de_sincronizar_y_apagar_sin_doble_worker(self):
        espera = self.fuente.split("def mostrar_ventana_espera_registro(", 1)[1].split(
            "def entregar_y_apagar", 1
        )[0]
        self.assertIn("solicitar_sincronizacion=False", espera)
        self.assertIn('"Registrando entrega..."', espera)
        self.assertIn("sincronizar_devolucion_pre_apagado(id_evento)", espera)
        self.assertNotIn("time.sleep(", espera)
        self.assertLess(
            espera.index("finalizar_devolucion_durable("),
            espera.index("sincronizar_devolucion_pre_apagado(id_evento)"),
        )
        self.assertLess(
            espera.index("sincronizar_devolucion_pre_apagado(id_evento)"),
            espera.index("apagar_windows()"),
        )
        self.assertIn("if exito:", espera)
        self.assertIn("if not sesion:\n            return False", espera)
        self.assertIn("state=\"disabled\"", self.fuente)

    def test_prestamo_bypass_reanuda_curp_y_devolucion_tiene_watchdog(self):
        prestamo = self.fuente.split("def iniciar_sesion", 1)[1].split(
            "def crear_pantalla_login", 1
        )[0]
        espera = self.fuente.split("def mostrar_ventana_espera_registro", 1)[1].split(
            "def entregar_y_apagar", 1
        )[0]
        self.assertIn('"PRESTAMO_NUEVO"', prestamo)
        self.assertIn("reanudar=lambda: procesar_proximidad(", prestamo)
        self.assertIn("CURP_SOLICITADA", self.fuente)
        self.assertIn("FLUJO_RECUPERACION_SESION", self.fuente)
        posicion_bypass = prestamo.index('"PRESTAMO_NUEVO"')
        self.assertNotEqual(
            prestamo.find("mostrar_confirmacion_simple(", posicion_bypass), -1
        )
        confirmacion = self.fuente.split("def mostrar_confirmacion_simple", 1)[1].split(
            "def cambiar_estado", 1
        )[0]
        self.assertIn("iniciar_validacion_identidad()", confirmacion)
        self.assertIn("finalizar_devolucion_durable(", espera)
        self.assertIn("DEVOLUCION_SYNC_INMEDIATA_INICIADA", espera)
        self.assertIn("DEVOLUCION_WATCHDOG_APAGADO", espera)
        self.assertIn("DEVOLUCION_WATCHDOG_RECUPERA_UI", espera)
        self.assertLess(
            espera.index("iniciar_watchdog()\n\n    # Ejecutar en un hilo separado"),
            espera.index("threading.Thread(target=ejecutar_registro, daemon=True).start()"),
        )
        self.assertIn('boton_entrega.config(state="normal", text="ENTREGAR Y APAGAR")', espera)
        self.assertIn("6000", espera)

    def test_bypass_admin_devolucion_reanuda_la_misma_entrega_durable(self):
        """LEJOS + CtrlAltU no puede depender del retorno de wait_window."""
        entrega = self.fuente.split("def entregar_y_apagar", 1)[1].split(
            "def mostrar_ventana_entrega", 1
        )[0]
        bypass = self.fuente.split("def _autorizar_bypass_proximidad", 1)[1].split(
            "def _solicitar_contingencia_proximidad", 1
        )[0]
        espera = self.fuente.split("def mostrar_ventana_espera_registro", 1)[1].split(
            "def entregar_y_apagar", 1
        )[0]

        self.assertIn('"DEVOLUCION"', entrega)
        self.assertIn("def iniciar_devolucion_durable(", entrega)
        self.assertIn("reanudar=reanudar_devolucion_administrativa", entrega)
        self.assertIn("iniciar_devolucion_durable(autorizacion_administrativa=True)", entrega)
        self.assertIn("DEVOLUCION_PROXIMIDAD_AUTORIZADA_ADMIN", entrega)
        self.assertIn("if not _widget_tk_vivo(ventana_ubicacion):", entrega)
        self.assertIn("DEVOLUCION_DURABLE_PIPELINE_INICIADO", entrega)
        self.assertIn('if contexto.get("reanudado"):', entrega)
        self.assertIn("parent.after_idle(reanudar)", bypass)
        self.assertIn("contexto[\"reanudado\"] = True", bypass)
        self.assertIn("DEVOLUCION_BYPASS_CONTINUACION_TIMEOUT", bypass)
        self.assertIn("finalizar_devolucion_durable(", espera)
        self.assertIn("sincronizar_devolucion_pre_apagado(id_evento)", espera)
        self.assertLess(
            espera.index("finalizar_devolucion_durable("),
            espera.index("apagar_windows()"),
        )

    def test_callback_bypass_reanuda_una_sola_devolucion_sin_wait_window(self):
        modulo = self._funciones_ast(
            "iniciar_contexto_contingencia", "_autorizar_bypass_proximidad"
        )
        eventos, usos, reanudaciones = [], [], []

        class ParentFalso:
            def __init__(self):
                self.pendientes = []

            def after_idle(self, callback):
                self.pendientes.append(callback)

        espacio = {
            "contexto_contingencia_proximidad": None,
            "registrar_evento_tecnico": eventos.append,
            "registrar_uso_contingencia": lambda operacion, matricula: usos.append(
                (operacion, matricula)
            ),
            "tk": SimpleNamespace(TclError=RuntimeError),
        }
        exec(compile(ast.fix_missing_locations(modulo), "<bypass>", "exec"), espacio)
        parent = ParentFalso()
        contexto = espacio["iniciar_contexto_contingencia"](
            "LEJOS", "DEVOLUCION", "0824118010", parent,
            reanudar=lambda: reanudaciones.append("DEVOLUCION_DURABLE"),
        )
        contexto["cerrar_dialogo_fallo"] = lambda: eventos.append("DIALOGO_CERRADO")

        espacio["_autorizar_bypass_proximidad"](contexto)
        espacio["_autorizar_bypass_proximidad"](contexto)

        self.assertTrue(contexto["autorizado"])
        self.assertTrue(contexto["reanudado"])
        self.assertEqual(reanudaciones, [])
        self.assertEqual(len(parent.pendientes), 1)
        parent.pendientes.pop()()
        self.assertEqual(reanudaciones, ["DEVOLUCION_DURABLE"])
        self.assertEqual(usos, [("DEVOLUCION", "0824118010")])

    def test_prestamo_online_espera_un_intento_acotado_fuera_de_tk(self):
        prestamo = self.fuente.split("def registrar_entrada(matricula):", 1)[1].split(
            "def _existe_entrada_del_mismo_intento", 1
        )[0]
        self.assertIn("registrar_prestamo_coordinador(", prestamo)
        self.assertIn("sincronizar_evento_inmediato(id_evento)", prestamo)
        self.assertNotIn("enviar_evento_coordinador(", prestamo)
        self.assertIn("threading.Thread(target=registrar_prestamo, daemon=True).start()", self.fuente)

    def test_matricula_del_entry_no_se_convierte_a_numero(self):
        inicio = self.fuente.split("def iniciar_sesion():", 1)[1].split(
            "def mostrar_pantalla_espera", 1
        )[0]
        self.assertIn("matricula = entrada.get().strip()", inicio)
        self.assertNotIn("int(matricula)", self.fuente)
        self.assertNotIn("zfill(", self.fuente)

    def _funciones_ast(self, *nombres):
        arbol = ast.parse(self.fuente.lstrip("\ufeff"))
        funciones = [
            nodo for nodo in arbol.body
            if isinstance(nodo, ast.FunctionDef) and nodo.name in nombres
        ]
        self.assertEqual({nodo.name for nodo in funciones}, set(nombres))
        return ast.Module(body=funciones, type_ignores=[])

    def test_ctrl_alt_u_reconoce_u_mayuscula_y_ambas_mascaras_alt(self):
        modulo = self._funciones_ast(
            "es_atajo_administrativo",
            "es_atajo_salida_administrativa",
        )
        espacio = {}
        exec(compile(ast.fix_missing_locations(modulo), "<atajo>", "exec"), espacio)

        detectar = espacio["es_atajo_administrativo"]
        for tecla in ("u", "U"):
            self.assertTrue(detectar(SimpleNamespace(keysym=tecla, state=0x0004 | 0x0008)))
            self.assertTrue(detectar(SimpleNamespace(keysym=tecla, state=0x0004 | 0x20000)))
        self.assertFalse(detectar(SimpleNamespace(keysym="u", state=0x0004)))
        self.assertFalse(detectar(SimpleNamespace(keysym="u", state=0x0008)))

    def test_root_y_cada_toplevel_usan_el_mismo_dispatcher_directo(self):
        self.assertIn("vincular_dispatcher_administrativo(ventana)", self.fuente)
        self.assertIn('ventana.bind_all("<KeyPress>", manejar_tecla_global)', self.fuente)
        vinculacion = self.fuente.split("def vincular_dispatcher_administrativo(widget):", 1)[1].split(
            "def manejar_tecla_global", 1
        )[0]
        self.assertIn('"<KeyPress>"', vinculacion)
        self.assertIn('"<Control-Alt-u>"', vinculacion)
        self.assertIn('"<Control-Alt-U>"', vinculacion)
        self.assertIn("manejar_evento_administrativo", vinculacion)

        # Cada constructor de Toplevel queda vinculado, también los que usan grab_set.
        self.assertGreaterEqual(
            self.fuente.count("vincular_dispatcher_administrativo(") - 1,
            self.fuente.count("tk.Toplevel("),
        )

    def test_binding_directo_es_idempotente_y_no_duplica_acciones(self):
        modulo = self._funciones_ast("vincular_dispatcher_administrativo")
        espacio = {"manejar_evento_administrativo": lambda *args, **kwargs: None}
        exec(compile(ast.fix_missing_locations(modulo), "<binding>", "exec"), espacio)

        class WidgetFalso:
            def __init__(self):
                self.bindings = []

            def bind(self, secuencia, callback, add=None):
                self.bindings.append((secuencia, callback, add))

        widget = WidgetFalso()
        espacio["vincular_dispatcher_administrativo"](widget)
        cantidad = len(widget.bindings)
        espacio["vincular_dispatcher_administrativo"](widget)
        self.assertEqual(len(widget.bindings), cantidad)
        self.assertEqual(widget.bindings[0][0], "<KeyPress>")
        self.assertTrue(all(binding[2] == "+" for binding in widget.bindings))

    def test_foco_en_boton_entry_o_toplevel_resuelve_su_ventana(self):
        contexto = self.fuente.split("def _contexto_visual_administrativo", 1)[1].split(
            "def _accion_atajo_administrativo", 1
        )[0]
        self.assertIn('getattr(event, "widget", None)', contexto)
        self.assertIn("candidato.winfo_toplevel()", contexto)
        self.assertIn("ventana.grab_current()", contexto)
        self.assertIn("ventana.focus_get()", contexto)

    def test_dispatcher_central_separa_contingencia_y_salida(self):
        modulo = self._funciones_ast("_accion_atajo_administrativo")
        espacio = {}
        exec(compile(ast.fix_missing_locations(modulo), "<accion>", "exec"), espacio)
        accion = espacio["_accion_atajo_administrativo"]
        for estado in ("NO_DETECTADO", "ERROR", "LEJOS"):
            self.assertEqual(accion(estado), "CONTINUAR_SIN_PROXIMIDAD")

    def test_atajo_central_envia_proximidad_a_bypass_y_salida_solo_shift(self):
        manejador = self.fuente.split(
            "def manejar_atajo_administrativo(", 1
        )[1].split(
            "def manejar_atajo_admin", 1
        )[0]
        self.assertIn("TECLA_ADMIN_RECIBIDA", manejador)
        self.assertIn("ATAJO_CTRL_ALT_U_DETECTADO", manejador)
        self.assertIn("ESTADO_PROXIMIDAD_ACTUAL", manejador)
        self.assertIn("CONTINGENCIA_ABIERTA_FLAG", manejador)
        self.assertIn("ATAJO_ADMIN_DETECTADO contexto=PROXIMIDAD", manejador)
        self.assertIn('accion == "CONTINUAR_SIN_PROXIMIDAD"', manejador)
        self.assertIn("ATAJO_CTRL_ALT_U_RECHAZADO motivo=SIN_PROXIMIDAD_ACTIVA", manejador)
        self.assertIn("_solicitar_salida_administrativa", manejador)
        self.assertIn("finally:", manejador)
        self.assertIn("atajo_administrativo_en_curso = False", manejador)

    def test_diagnostico_de_tecla_no_registra_texto_escrito(self):
        diagnostico = self.fuente.split("def registrar_tecla_proximidad(event):", 1)[1].split(
            "def _contexto_visual_administrativo", 1
        )[0]
        for campo in ("TECLA_PROXIMIDAD", "keysym=", "keycode=", "state=", "widget="):
            self.assertIn(campo, diagnostico)
        self.assertNotIn("event.char", diagnostico)
        self.assertIn('widget.winfo_class()', diagnostico)

    def test_estados_de_proximidad_tienen_texto_y_log_inequivocos(self):
        self.assertEqual(self.fuente.count("iniciar_contexto_contingencia("), 3)
        self.assertEqual(
            self.fuente.count("al_crear=registrar_dialogo_fallo_contingencia"),
            2,
        )
        self.assertIn('"PROXIMIDAD_DIALOGO_ABIERTO estado={estado}"', self.fuente)
        self.assertIn('"Carrito no detectado"', self.fuente)
        self.assertIn('"Error de verificación"', self.fuente)
        self.assertIn('"Laptop fuera de la ubicación autorizada"', self.fuente)
        self.assertNotIn('"Ubicación no validada"', self.fuente)

    def test_escape_admin_independiente_existe_en_todos_los_estados(self):
        vinculacion = self.fuente.split("def vincular_dispatcher_administrativo", 1)[1].split(
            "def manejar_tecla_global", 1
        )[0]
        self.assertIn('"<Control-Alt-Shift-u>"', vinculacion)
        self.assertIn('"<Control-Alt-Shift-U>"', vinculacion)
        self.assertIn("forzar_salida=True", vinculacion)
        self.assertIn("if contexto_proximidad and not forzar_salida:", self.fuente)

    def test_admin_libera_y_restaura_grab_modal_al_cancelar(self):
        admin = self.fuente.split("def abrir_acceso_administrativo", 1)[1].split(
            "def iniciar_contexto_contingencia", 1
        )[0]
        self.assertIn("ventana.grab_current()", admin)
        self.assertIn("grab_anterior.grab_release()", admin)
        self.assertIn("ADMIN_GRAB_TEMPORAL_LIBERADO", admin)
        self.assertIn("_restaurar_grab_administrativo", admin)
        self.assertIn("ADMIN_SOLICITUD estado_guard=", admin)
        self.assertIn("ADMIN_DIALOGO_ABIERTO", admin)
        self.assertIn("entrada_clave.focus_force()", admin)

    def test_cleanup_admin_es_unico_idempotente_y_timeout_lo_invoca(self):
        cleanup = self.fuente.split("def finalizar_dialogo_administrativo", 1)[1].split(
            "def abrir_acceso_administrativo", 1
        )[0]
        for traza in (
            "ADMIN_DIALOGO_CIERRE motivo=",
            "ADMIN_ESTADO_LIMPIADO",
            "ADMIN_REAPERTURA_PERMITIDA",
        ):
            self.assertIn(traza, cleanup)
        self.assertIn("after_cancel(after_id)", cleanup)
        self.assertIn("dialogo_admin_abierto = False", cleanup)
        self.assertIn("dialogo_admin_referencia = None", cleanup)
        self.assertIn("dialogo_admin_estado = None", cleanup)
        self.assertIn('finalizar_dialogo_administrativo("TIMEOUT", restaurar=False)', self.fuente)

    def test_timeout_admin_limpia_guard_y_permite_tres_reaperturas(self):
        modulo = self._funciones_ast(
            "_widget_tk_vivo",
            "finalizar_dialogo_administrativo",
        )

        class TclErrorFalso(Exception):
            pass

        logs = []
        espacio = {
            "tk": SimpleNamespace(TclError=TclErrorFalso),
            "registrar_evento_tecnico": logs.append,
            "_restaurar_grab_administrativo": lambda *_args: None,
            "dialogo_admin_abierto": False,
            "dialogo_admin_referencia": None,
            "dialogo_admin_estado": None,
        }
        exec(compile(ast.fix_missing_locations(modulo), "<admin>", "exec"), espacio)

        class DialogoFalso:
            def __init__(self):
                self.vivo = True
                self.cancelados = []

            def winfo_exists(self):
                return self.vivo

            def after_cancel(self, after_id):
                self.cancelados.append(after_id)

            def grab_release(self):
                return None

            def destroy(self):
                self.vivo = False

        for _ in range(3):
            dialogo = DialogoFalso()
            espacio["dialogo_admin_abierto"] = True
            espacio["dialogo_admin_referencia"] = dialogo
            espacio["dialogo_admin_estado"] = {
                "dialogo": dialogo,
                "grab_anterior": None,
                "foco_anterior": None,
                "after_pendientes": ["after-prueba"],
                "finalizado": False,
            }
            self.assertTrue(
                espacio["finalizar_dialogo_administrativo"]("TIMEOUT", restaurar=False)
            )
            self.assertFalse(espacio["dialogo_admin_abierto"])
            self.assertIsNone(espacio["dialogo_admin_referencia"])
            self.assertIsNone(espacio["dialogo_admin_estado"])
            self.assertEqual(dialogo.cancelados, ["after-prueba"])
            self.assertFalse(
                espacio["finalizar_dialogo_administrativo"]("TIMEOUT", restaurar=False)
            )

        self.assertEqual(logs.count("ADMIN_REAPERTURA_PERMITIDA"), 3)

    def test_todos_los_cierres_admin_pasan_por_cleanup_central(self):
        admin = self.fuente.split("def abrir_acceso_administrativo", 1)[1].split(
            "def iniciar_contexto_contingencia", 1
        )[0]
        self.assertIn('cerrar("AUTORIZADO", restaurar=False)', admin)
        self.assertIn('cerrar("PASSWORD_INCORRECTA")', admin)
        self.assertIn('lambda: cerrar("CANCELADO")', admin)
        self.assertIn('lambda _evento: cerrar("CANCELADO")', admin)
        self.assertIn('finalizar_dialogo_administrativo("DESTROY_PROGRAMATICO", dialogo=dialogo)', admin)
        self.assertIn('dialogo.bind("<Destroy>", al_destruir, add="+")', admin)

    def test_restaurar_grab_tolera_modal_previo_destruido(self):
        modulo = self._funciones_ast("_widget_tk_vivo", "_restaurar_grab_administrativo")

        class TclErrorFalso(Exception):
            pass

        logs = []
        espacio = {
            "tk": SimpleNamespace(TclError=TclErrorFalso),
            "registrar_evento_tecnico": logs.append,
        }
        exec(compile(ast.fix_missing_locations(modulo), "<grab>", "exec"), espacio)

        class WidgetFalso:
            def __init__(self, vivo=True):
                self.vivo = vivo
                self.grab_restaurado = False
                self.foco = False

            def winfo_exists(self):
                return self.vivo

            def grab_set(self):
                self.grab_restaurado = True

            def focus_force(self):
                self.foco = True

        destruido = WidgetFalso(vivo=False)
        espacio["_restaurar_grab_administrativo"](destruido, None)
        self.assertFalse(destruido.grab_restaurado)

        vivo = WidgetFalso()
        espacio["_restaurar_grab_administrativo"](vivo, None)
        self.assertTrue(vivo.grab_restaurado)
        self.assertTrue(vivo.foco)
        self.assertIn("ADMIN_GRAB_RESTAURADO", logs)

    def test_dialogo_admin_separa_bypass_y_salida_por_intencion(self):
        admin = self.fuente.split("def abrir_acceso_administrativo", 1)[1].split(
            "def iniciar_contexto_contingencia", 1
        )[0]
        self.assertIn("hmac.compare_digest(hash_ingresado, HASH_CLAVE_ADMIN)", admin)
        self.assertIn('accion="SALIR_SISTEMA"', admin)
        self.assertIn('"CONTINUAR_SIN_PROXIMIDAD"', admin)
        self.assertIn("ADMIN_DIALOGO_SOLICITADO accion={accion}", admin)
        self.assertIn("ADMIN_AUTORIZACION_CORRECTA accion={accion}", admin)
        self.assertIn('if accion == "CONTINUAR_SIN_PROXIMIDAD":', admin)
        self.assertIn("al_autorizar()", admin)
        self.assertIn("SALIDA_ADMIN_AUTORIZADA", admin)
        self.assertIn("cerrar_aplicacion()", admin)
        self.assertIn('show="*"', admin)
        self.assertNotIn("registrar_evento_tecnico(clave", admin)

    def test_bypass_autorizado_cierra_solo_dialogo_de_proximidad(self):
        modulo = self._funciones_ast("_autorizar_bypass_proximidad")
        eventos = []
        usos = []
        espacio = {
            "contexto_contingencia_proximidad": None,
            "registrar_evento_tecnico": eventos.append,
            "registrar_uso_contingencia": lambda operacion, matricula: usos.append(
                (operacion, matricula)
            ),
        }
        exec(compile(ast.fix_missing_locations(modulo), "<bypass>", "exec"), espacio)

        for estado in ("LEJOS", "NO_DETECTADO", "ERROR"):
            cierres = []
            contexto = {
                "estado": estado,
                "operacion": "PRESTAMO",
                "matricula": "0824118010",
                "autorizado": False,
                "cerrar_dialogo_fallo": lambda: cierres.append(estado),
            }
            espacio["contexto_contingencia_proximidad"] = contexto
            espacio["_autorizar_bypass_proximidad"](contexto)
            self.assertTrue(contexto["autorizado"])
            self.assertEqual(cierres, [estado])
            espacio["_autorizar_bypass_proximidad"](contexto)
            self.assertEqual(cierres, [estado])

        self.assertEqual(len(usos), 3)
        self.assertIn("PROXIMIDAD_BYPASS_ADMIN_AUTORIZADO estado=LEJOS", eventos)

    def test_bypass_reanuda_prestamo_y_devolucion_sin_duplicar_sesion(self):
        self.assertEqual(self.fuente.count("if contingencia_autorizada:"), 2)
        self.assertEqual(self.fuente.count("procesar_resultado_proximidad(True, None)"), 1)
        self.assertIn("reanudar=lambda: procesar_proximidad(", self.fuente)
        self.assertIn("mostrar_confirmacion_simple(", self.fuente)
        entrega = self.fuente.split("def entregar_y_apagar", 1)[1].split(
            "def mostrar_ventana_entrega", 1
        )[0]
        ruta_durable = entrega.split("def iniciar_devolucion_durable", 1)[1].split(
            "def comprobar_devolucion", 1
        )[0]
        self.assertIn("mostrar_ventana_espera_registro(", ruta_durable)
        self.assertNotIn("verificar_sheets_con_reintentos", ruta_durable)
        solicitud = self.fuente.split("def _solicitar_contingencia_proximidad", 1)[1].split(
            "def _solicitar_salida_administrativa", 1
        )[0]
        self.assertIn('accion="CONTINUAR_SIN_PROXIMIDAD"', solicitud)
        self.assertIn("al_autorizar=lambda: _autorizar_bypass_proximidad(contexto)", solicitud)

    def test_bypass_devolucion_sin_callback_recupera_la_ui(self):
        modulo = self._funciones_ast(
            "iniciar_contexto_contingencia", "_autorizar_bypass_proximidad"
        )
        eventos, recuperaciones = [], []

        class ParentFalso:
            def __init__(self):
                self.pendientes = []
                self.guards = []

            def after_idle(self, callback):
                self.pendientes.append(callback)

            def after(self, _milisegundos, callback):
                self.guards.append(callback)
                return "guard-1"

        espacio = {
            "contexto_contingencia_proximidad": None,
            "registrar_evento_tecnico": eventos.append,
            "registrar_uso_contingencia": lambda *_: None,
            "tk": SimpleNamespace(TclError=RuntimeError),
        }
        exec(compile(ast.fix_missing_locations(modulo), "<bypass>", "exec"), espacio)
        parent = ParentFalso()
        contexto = espacio["iniciar_contexto_contingencia"](
            "LEJOS", "DEVOLUCION", "A1", parent,
            reanudar=lambda: None,
            recuperar_bypass=lambda: recuperaciones.append("UI_RESTAURADA"),
        )
        espacio["_autorizar_bypass_proximidad"](contexto)
        self.assertEqual(len(parent.guards), 1)
        parent.guards[0]()
        self.assertEqual(recuperaciones, ["UI_RESTAURADA"])
        self.assertIn(
            "DEVOLUCION_BYPASS_CONTINUACION_TIMEOUT "
            "motivo=CALLBACK_NO_INICIO_DURABLE", eventos
        )

    def test_salida_shift_usa_intencion_exclusiva(self):
        salida = self.fuente.split("def _solicitar_salida_administrativa", 1)[1].split(
            "def manejar_evento_administrativo", 1
        )[0]
        self.assertIn('accion="SALIR_SISTEMA"', salida)
        self.assertIn('"<Control-Alt-Shift-u>"', self.fuente)
        self.assertIn('"<Control-Alt-Shift-U>"', self.fuente)

    def test_cancelar_o_clave_incorrecta_no_ejecutan_el_callback_bypass(self):
        admin = self.fuente.split("def abrir_acceso_administrativo", 1)[1].split(
            "def iniciar_contexto_contingencia", 1
        )[0]
        autenticar = admin.split("    def autenticar():", 1)[1].split(
            "    tk.Button(", 1
        )[0]
        self.assertIn("completar_autorizacion()", autenticar)
        self.assertIn("return", autenticar)
        self.assertIn('intentos_incorrectos["cantidad"] += 1', autenticar)
        self.assertNotIn("al_autorizar", autenticar)
        self.assertIn("command=cerrar", admin)

    def test_clave_incorrecta_no_cierra_y_flags_se_liberan_ante_excepcion(self):
        admin = self.fuente.split("def abrir_acceso_administrativo", 1)[1].split(
            "def iniciar_contexto_contingencia", 1
        )[0]
        self.assertIn('estado_intentos.config(', admin)
        self.assertIn("if intentos_incorrectos[\"cantidad\"] >= 3:", admin)
        manejador = self.fuente.split("def manejar_atajo_administrativo(", 1)[1].split(
            "def manejar_atajo_admin", 1
        )[0]
        self.assertIn("try:", manejador)
        self.assertIn("finally:", manejador)
        self.assertIn("atajo_administrativo_en_curso = False", manejador)

    def test_antirrebote_y_flag_impiden_doble_disparo(self):
        manejador = self.fuente.split("def manejar_atajo_administrativo(", 1)[1].split(
            "def manejar_atajo_admin", 1
        )[0]
        self.assertIn("atajo_administrativo_en_curso", manejador)
        self.assertIn("ultimo_atajo_administrativo", manejador)
        self.assertIn("ATAJO_ADMIN_OMITIDO motivo=doble_disparo", manejador)
        self.assertIn("ahora - ultimo_atajo_administrativo < 0.20", manejador)

    def test_cerrar_y_reabrir_proximidad_vuelve_a_vincular_sin_estado_global(self):
        registro = self.fuente.split("def registrar_dialogo_fallo_contingencia", 1)[1].split(
            "def finalizar_contexto_contingencia", 1
        )[0]
        self.assertIn("vincular_dispatcher_administrativo(_dialogo)", registro)
        self.assertIn('contexto_contingencia_proximidad["dialogo"] = _dialogo', registro)


if __name__ == "__main__":
    unittest.main()
