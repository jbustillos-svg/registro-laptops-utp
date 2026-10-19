from pathlib import Path
import sqlite3
import subprocess
import tempfile
import unittest
import json

from contingencia_local import (
    CACHE_NO_DISPONIBLE,
    CACHE_OK,
    SCHEMA_VERSION_ACTUAL,
    ContingenciaLocal,
)


class PruebasContingenciaLocal(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.almacen = ContingenciaLocal(Path(self.temporal.name))

    def tearDown(self):
        self.temporal.cleanup()

    def test_primera_ejecucion_crea_base_y_esquema(self):
        self.assertFalse(self.almacen.ruta_db.exists())
        self.almacen.inicializar()
        self.assertTrue(self.almacen.ruta_db.exists())
        self.assertEqual(self.almacen.version_esquema(), SCHEMA_VERSION_ACTUAL)
        self.assertEqual(self.almacen.estado_cache(), CACHE_NO_DISPONIBLE)

    def test_migraciones_repetidas_conservan_registros(self):
        self.almacen.inicializar()
        self.almacen.upsert_persona("A1", nombre="Ana", rol="ALUMNO")
        evento = self.almacen.crear_evento("E1", "PRESTAMO")
        self.almacen.inicializar()
        self.assertEqual(self.almacen.obtener_persona("A1")["nombre"], "Ana")
        self.assertIsNotNone(self.almacen.obtener_evento(evento))

    def test_migra_base_v1_sin_borrar_eventos_pendientes(self):
        ruta = self.almacen.ruta_db
        conexion = sqlite3.connect(ruta)
        conexion.executescript(
            """
            CREATE TABLE schema_version (id INTEGER PRIMARY KEY, version INTEGER NOT NULL);
            INSERT INTO schema_version VALUES (1, 1);
            CREATE TABLE personas_cache (
                matricula TEXT PRIMARY KEY, nombre TEXT NOT NULL DEFAULT '',
                rol TEXT NOT NULL DEFAULT 'ALUMNO', curp_ultimos_2 TEXT, estado TEXT,
                no_entregas INTEGER, grupo TEXT, carrera TEXT, fecha_actualizacion TEXT NOT NULL
            );
            CREATE TABLE eventos_pendientes (
                id_evento TEXT PRIMARY KEY, equipo TEXT, tipo_evento TEXT NOT NULL,
                matricula TEXT, nombre TEXT, fecha_evento TEXT NOT NULL, hora_evento TEXT NOT NULL,
                laptop_id TEXT, datos_json TEXT NOT NULL DEFAULT '{}',
                sincronizado INTEGER NOT NULL DEFAULT 0, intentos INTEGER NOT NULL DEFAULT 0,
                ultimo_error TEXT, fecha_creacion TEXT NOT NULL
            );
            INSERT INTO eventos_pendientes VALUES (
                'pendiente-v1', 'LAP-1', 'PRESTAMO', 'A1', 'Ana', '2026-10-16',
                '10:00:00', 'LAP-1', '{}', 0, 0, NULL, '2026-10-16T10:00:00-06:00'
            );
            """
        )
        conexion.commit()
        conexion.close()
        self.almacen.inicializar()
        self.assertEqual(self.almacen.version_esquema(), SCHEMA_VERSION_ACTUAL)
        self.assertIsNotNone(self.almacen.obtener_evento("pendiente-v1"))

    def test_upsert_y_lectura_de_persona(self):
        self.almacen.inicializar()
        self.almacen.upsert_persona("A1", "Ana", "ALUMNO", "XY", "ACTIVO", 1)
        persona = self.almacen.upsert_persona("A1", nombre="Ana López", no_entregas=2)
        self.assertEqual(persona["nombre"], "Ana López")
        self.assertEqual(persona["rol"], "ALUMNO")
        self.assertEqual(persona["curp_ultimos_2"], "XY")
        self.assertEqual(persona["no_entregas"], 2)
        self.assertEqual(self.almacen.estado_cache(), CACHE_OK)

    def test_fallo_externo_no_reemplaza_cache_valida(self):
        self.almacen.inicializar()
        self.almacen.upsert_persona("A1", nombre="Ana", rol="ALUMNO")
        # Un fallo de Sheets no llama a upsert; la caché persistente sigue intacta.
        reiniciado = ContingenciaLocal(self.temporal.name)
        reiniciado.inicializar()
        self.assertEqual(reiniciado.obtener_persona("A1")["nombre"], "Ana")

    def test_eventos_idempotentes_y_preservan_hora_original(self):
        self.almacen.inicializar()
        evento = self.almacen.crear_evento(
            "E1", "PRESTAMO", "A1", "Ana", "2026-10-16", "10:14:32",
            id_evento="evento-fijo",
        )
        self.almacen.crear_evento(
            "E1", "PRESTAMO", "A1", "Cambiado", "2026-10-16", "11:00:00",
            id_evento=evento,
        )
        self.assertEqual(len(self.almacen.obtener_eventos_pendientes()), 1)
        guardado = self.almacen.obtener_evento(evento)
        self.assertEqual((guardado["fecha_evento"], guardado["hora_evento"]),
                         ("2026-10-16", "10:14:32"))

    def test_eventos_persisten_intentos_error_y_sincronizacion(self):
        self.almacen.inicializar()
        evento = self.almacen.crear_evento("E1", "DEVOLUCION")
        actualizado = self.almacen.incrementar_intentos(evento, "timeout")
        self.assertEqual((actualizado["intentos"], actualizado["ultimo_error"]), (1, "timeout"))
        self.almacen.guardar_ultimo_error(evento, "api")
        reiniciado = ContingenciaLocal(self.temporal.name)
        self.assertEqual(reiniciado.obtener_evento(evento)["ultimo_error"], "api")
        reiniciado.marcar_sincronizado(evento)
        self.assertEqual(reiniciado.obtener_eventos_pendientes(), [])

    def test_prestamo_offline_es_idempotente_y_sobrevive_reinicio(self):
        self.almacen.inicializar()
        primera = self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "10:14:32"
        )
        segunda = self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "10:14:32"
        )
        self.assertEqual(primera["id_sesion"], segunda["id_sesion"])
        self.assertEqual(len(self.almacen.obtener_eventos_pendientes()), 1)
        reiniciado = ContingenciaLocal(self.temporal.name)
        sesion = reiniciado.obtener_sesion_activa("LAP-1")
        self.assertEqual((sesion["matricula"], sesion["hora_entrada"]), ("A1", "10:14:32"))
        self.assertFalse(sesion["prestamo_sincronizado"])

    def test_devolucion_offline_conserva_hora_y_no_duplica_evento(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "10:14:32"
        )
        primera = self.almacen.registrar_devolucion_offline(
            "A1", "LAP-1", "2026-10-16", "11:19:40"
        )
        segunda = self.almacen.registrar_devolucion_offline(
            "A1", "LAP-1", "2026-10-16", "12:00:00"
        )
        self.assertEqual((primera["hora_salida"], segunda["hora_salida"]), ("11:19:40", "11:19:40"))
        eventos = self.almacen.obtener_eventos_pendientes()
        self.assertEqual({evento["tipo_evento"] for evento in eventos}, {"PRESTAMO", "DEVOLUCION"})
        devolucion = next(evento for evento in eventos if evento["tipo_evento"] == "DEVOLUCION")
        self.assertEqual(devolucion["hora_evento"], "11:19:40")

    def test_prestamo_online_admite_devolucion_offline(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_online(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        self.almacen.registrar_devolucion_offline(
            "A1", "LAP-1", "2026-10-16", "10:00:00"
        )
        eventos = self.almacen.obtener_eventos_pendientes()
        self.assertEqual(len(eventos), 1)
        self.assertEqual((eventos[0]["tipo_evento"], eventos[0]["hora_evento"]),
                         ("DEVOLUCION", "10:00:00"))

    def test_prestamo_coordinador_tiene_uuid_estable_y_pendiente(self):
        self.almacen.inicializar()
        primero = self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        segundo = self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        self.assertEqual(primero["prestamo_evento_id"], segundo["prestamo_evento_id"])
        evento = self.almacen.obtener_evento(primero["prestamo_evento_id"])
        self.assertEqual(evento["tipo_evento"], "PRESTAMO")
        self.assertEqual(evento["sincronizado"], 0)

    def test_sesion_previa_cerrada_permite_nuevo_prestamo_y_evento_unico(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        self.almacen.registrar_devolucion_online("A1", "LAP-1", "10:00:00")
        nuevo = self.almacen.registrar_prestamo_coordinador(
            "B1", "Beto", "LAP-1", "2026-10-16", "11:00:00"
        )
        repetido = self.almacen.registrar_prestamo_coordinador(
            "B1", "Beto", "LAP-1", "2026-10-16", "11:00:00"
        )
        self.assertFalse(nuevo["sesion_existente"])
        self.assertTrue(repetido["sesion_existente"])
        self.assertEqual(nuevo["prestamo_evento_id"], repetido["prestamo_evento_id"])
        self.assertEqual(
            len([evento for evento in self.almacen.obtener_eventos_pendientes()
                 if evento["tipo_evento"] == "PRESTAMO" and evento["matricula"] == "B1"]),
            1,
        )

    def test_reconcilia_cierre_remoto_aplicado_antes_de_nuevo_prestamo(self):
        self.almacen.inicializar()
        anterior = self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        cerrada = self.almacen.registrar_devolucion_offline(
            "A1", "LAP-1", "2026-10-16", "10:00:00"
        )
        self.almacen.marcar_sincronizado(cerrada["devolucion_evento_id"], "Registros!2")
        # Simula una interrupción antigua: el cierre remoto está aplicado, la fila local quedó activa.
        with self.almacen._conexion() as conexion:
            conexion.execute(
                "UPDATE sesiones_locales SET hora_salida = NULL WHERE id_sesion = ?",
                (anterior["id_sesion"],),
            )
        nuevo = self.almacen.registrar_prestamo_coordinador(
            "B1", "Beto", "LAP-1", "2026-10-16", "11:00:00"
        )
        self.assertTrue(nuevo["sesion_reconciliada"])
        self.assertEqual(nuevo["matricula"], "B1")

    def test_sesion_realmente_activa_bloquea_otro_prestamo(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        with self.assertRaisesRegex(RuntimeError, "sesión activa"):
            self.almacen.registrar_prestamo_coordinador(
                "B1", "Beto", "LAP-1", "2026-10-16", "11:00:00"
            )

    def test_sesion_cerrada_tras_reinicio_no_se_recupera_para_devolucion(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_online(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        self.almacen.registrar_devolucion_online("A1", "LAP-1", "10:00:00")

        reiniciado = ContingenciaLocal(self.temporal.name)
        self.assertIsNone(reiniciado.obtener_sesion_activa("LAP-1"))

    def test_reinicio_con_sesion_activa_no_genera_no_entrega(self):
        self.almacen.inicializar()
        creada = self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )

        reiniciado = ContingenciaLocal(self.temporal.name)
        pendiente = reiniciado.obtener_sesion_activa("LAP-1")
        self.assertEqual(pendiente["id_sesion"], creada["id_sesion"])
        self.assertEqual(
            [evento["tipo_evento"] for evento in reiniciado.obtener_eventos_pendientes()],
            ["PRESTAMO"],
        )

    def test_misma_matricula_tras_reinicio_reutiliza_sesion_sin_nuevo_prestamo(self):
        self.almacen.inicializar()
        creada = self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )

        reiniciado = ContingenciaLocal(self.temporal.name)
        resultado = reiniciado.resolver_sesion_activa_por_nueva_matricula(
            "LAP-1", "A1", "2026-10-16", "10:00:00"
        )
        repetida = reiniciado.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "10:00:00"
        )

        self.assertEqual(resultado["resultado"], "MISMA_MATRICULA")
        self.assertEqual(resultado["sesion"]["id_sesion"], creada["id_sesion"])
        self.assertTrue(repetida["sesion_existente"])
        self.assertEqual(repetida["id_sesion"], creada["id_sesion"])
        self.assertEqual(
            len([e for e in reiniciado.obtener_eventos_pendientes()
                 if e["tipo_evento"] == "PRESTAMO"]),
            1,
        )

    def test_matricula_diferente_cierra_una_vez_y_permite_prestamo_offline(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )

        reiniciado = ContingenciaLocal(self.temporal.name)
        cerrado = reiniciado.resolver_sesion_activa_por_nueva_matricula(
            "LAP-1", "B1", "2026-10-16", "10:00:00"
        )
        reintento = reiniciado.resolver_sesion_activa_por_nueva_matricula(
            "LAP-1", "B1", "2026-10-16", "10:01:00"
        )
        nuevo = reiniciado.registrar_prestamo_offline(
            "B1", "Beto", "LAP-1", "2026-10-16", "10:02:00"
        )

        self.assertEqual(cerrado["resultado"], "SESION_ANTERIOR_CERRADA")
        self.assertEqual(cerrado["sesion"]["matricula"], "A1")
        self.assertEqual(cerrado["sesion"]["observacion"], "CIERRE_AUTOMATICO_POR_NUEVA_SESION")
        self.assertEqual(reintento["resultado"], "SIN_SESION")
        eventos = reiniciado.obtener_eventos_pendientes()
        self.assertEqual(
            len([e for e in eventos if e["tipo_evento"] == "NO_ENTREGA"]), 1
        )
        self.assertEqual(
            len([e for e in eventos if e["tipo_evento"] == "CIERRE_AUTOMATICO"]), 1
        )
        self.assertEqual(nuevo["matricula"], "B1")
        self.assertEqual(reiniciado.obtener_sesion_activa("LAP-1")["matricula"], "B1")

    def test_sincronizacion_tardia_y_reintentos_de_a_no_cierran_sesion_b(self):
        self.almacen.inicializar()
        sesion_a = self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        cierre_a = self.almacen.resolver_sesion_activa_por_nueva_matricula(
            "LAP-1", "B1", "2026-10-16", "10:00:00"
        )["sesion"]
        sesion_b = self.almacen.registrar_prestamo_coordinador(
            "B1", "Beto", "LAP-1", "2026-10-16", "10:01:00"
        )

        for campo in ("no_entrega_evento_id", "cierre_automatico_evento_id"):
            evento_a = self.almacen.obtener_evento(cierre_a[campo])
            self.assertEqual(json.loads(evento_a["datos_json"])["id_sesion"], sesion_a["id_sesion"])
            self.almacen.marcar_sincronizado(evento_a["id_evento"], "Registros!2")

        for _ in range(3):
            self.almacen.registrar_no_entrega_local(
                "A1", "LAP-1", "2026-10-16", "10:02:00",
                id_sesion=sesion_a["id_sesion"],
            )
            self.almacen.registrar_cierre_automatico_local(
                "A1", "LAP-1", "2026-10-16", "10:02:00",
                id_sesion=sesion_a["id_sesion"],
            )

        activa = self.almacen.obtener_sesion_activa("LAP-1")
        self.assertEqual(activa["id_sesion"], sesion_b["id_sesion"])
        self.assertEqual(activa["matricula"], "B1")

        devolucion_b = self.almacen.registrar_devolucion_offline(
            "B1", "LAP-1", "2026-10-16", "10:03:00",
            id_sesion=sesion_b["id_sesion"],
        )
        self.assertEqual(devolucion_b["id_sesion"], sesion_b["id_sesion"])
        self.assertIsNone(self.almacen.obtener_sesion_activa("LAP-1"))

    def test_cierre_por_id_sesion_inexistente_no_toca_sesion_activa(self):
        self.almacen.inicializar()
        sesion_b = self.almacen.registrar_prestamo_offline(
            "B1", "Beto", "LAP-1", "2026-10-16", "10:00:00"
        )

        with self.assertRaisesRegex(RuntimeError, "sesi.n local para cerrar"):
            self.almacen.registrar_cierre_automatico_local(
                "A1", "LAP-1", "2026-10-16", "10:01:00",
                id_sesion="sesion-inexistente",
            )

        activa = self.almacen.obtener_sesion_activa("LAP-1")
        self.assertEqual(activa["id_sesion"], sesion_b["id_sesion"])
        self.assertIsNone(activa["hora_salida"])

    def test_devolucion_tras_reinicio_reutiliza_la_sesion_activa_correcta(self):
        self.almacen.inicializar()
        creada = self.almacen.registrar_prestamo_coordinador(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )

        reiniciado = ContingenciaLocal(self.temporal.name)
        recuperada = reiniciado.obtener_sesion_activa("LAP-1")
        self.assertEqual(recuperada["id_sesion"], creada["id_sesion"])
        self.assertEqual((recuperada["matricula"], recuperada["nombre"]), ("A1", "Ana"))

        cerrada = reiniciado.registrar_devolucion_offline(
            "A1", "LAP-1", "2026-10-16", "10:00:00"
        )
        self.assertEqual(cerrada["id_sesion"], creada["id_sesion"])
        self.assertEqual(cerrada["hora_salida"], "10:00:00")
        self.assertIsNone(reiniciado.obtener_sesion_activa("LAP-1"))

    def test_error_sqlite_revierte_sesion_y_evento_de_prestamo(self):
        self.almacen.inicializar()
        with self.almacen._conexion() as conexion:
            conexion.execute(
                """
                CREATE TRIGGER bloquear_prestamo_prueba
                BEFORE INSERT ON sesiones_locales
                BEGIN SELECT RAISE(ABORT, 'FALLO_SQLITE_PRUEBA'); END
                """
            )
        with self.assertRaisesRegex(sqlite3.IntegrityError, "FALLO_SQLITE_PRUEBA"):
            self.almacen.registrar_prestamo_coordinador(
                "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
            )
        self.assertIsNone(self.almacen.obtener_sesion_activa("LAP-1"))
        self.assertEqual(self.almacen.obtener_eventos_pendientes(), [])

    def test_devolucion_online_cierra_sesion_sin_evento_pendiente(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_online(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        self.almacen.registrar_devolucion_online("A1", "LAP-1", "10:00:00")
        self.assertIsNone(self.almacen.obtener_sesion_activa("LAP-1"))
        self.assertEqual(self.almacen.obtener_eventos_pendientes(), [])

    def test_laptop_puede_tener_nueva_sesion_despues_de_cerrar_anterior(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_online(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        self.almacen.registrar_devolucion_online("A1", "LAP-1", "10:00:00")
        siguiente = self.almacen.registrar_prestamo_offline(
            "B1", "Beto", "LAP-1", "2026-10-16", "11:00:00"
        )
        self.assertEqual(siguiente["matricula"], "B1")
        self.assertEqual(self.almacen.obtener_sesion_activa("LAP-1")["matricula"], "B1")

    def test_no_entrega_y_cierre_automatico_son_idempotentes(self):
        self.almacen.inicializar()
        self.almacen.registrar_prestamo_offline(
            "A1", "Ana", "LAP-1", "2026-10-16", "09:00:00"
        )
        self.almacen.registrar_no_entrega_local(
            "A1", "LAP-1", "2026-10-16", "18:00:00"
        )
        self.almacen.registrar_no_entrega_local(
            "A1", "LAP-1", "2026-10-16", "18:30:00"
        )
        self.assertEqual(self.almacen.no_entregas_efectivas("A1", 2), 3)
        self.assertEqual(
            len([e for e in self.almacen.obtener_eventos_pendientes() if e["tipo_evento"] == "NO_ENTREGA"]),
            1,
        )
        self.almacen.registrar_prestamo_offline(
            "B1", "Beto", "LAP-2", "2026-10-16", "09:00:00"
        )
        sesion = self.almacen.registrar_cierre_automatico_local(
            "B1", "LAP-2", "2026-10-16", "10:00:00"
        )
        self.assertEqual(sesion["observacion"], "CIERRE_AUTOMATICO_POR_NUEVA_SESION")

    def test_archivos_sqlite_locales_no_se_trackean(self):
        raiz = Path(__file__).parent
        for nombre in (
            "contingencia.db", "contingencia.db-wal", "contingencia.db-shm",
            "contingencia.db-journal",
        ):
            resultado = subprocess.run(
                ["git", "check-ignore", "-q", nombre], cwd=raiz, check=False
            )
            self.assertEqual(resultado.returncode, 0, nombre)


if __name__ == "__main__":
    unittest.main()
