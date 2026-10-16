from pathlib import Path
import sqlite3
import subprocess
import tempfile
import unittest

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
