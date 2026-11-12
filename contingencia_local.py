"""Almacenamiento local duradero para la futura contingencia offline.

Este módulo no mantiene conexiones globales: cada operación abre, configura y
cierra su propia conexión SQLite. Así puede usarse desde los workers actuales
sin compartir una conexión entre hilos.
"""

from contextlib import contextmanager
from datetime import datetime
import json
from pathlib import Path
import sqlite3
import uuid


SCHEMA_VERSION_ACTUAL = 6

INTERNET_OK = "INTERNET_OK"
INTERNET_ERROR = "INTERNET_ERROR"
SHEETS_OK = "SHEETS_OK"
SHEETS_TIMEOUT = "SHEETS_TIMEOUT"
SHEETS_AUTH_ERROR = "SHEETS_AUTH_ERROR"
SHEETS_API_ERROR = "SHEETS_API_ERROR"
CACHE_OK = "CACHE_OK"
CACHE_NO_DISPONIBLE = "CACHE_NO_DISPONIBLE"


class ContingenciaLocal:
    """Acceso transaccional a ``contingencia.db`` dentro de la app instalada."""

    def __init__(self, directorio_app):
        self.directorio_app = Path(directorio_app)
        self.ruta_db = self.directorio_app / "contingencia.db"

    @contextmanager
    def _conexion(self):
        conexion = sqlite3.connect(
            str(self.ruta_db), timeout=10, isolation_level=None
        )
        conexion.row_factory = sqlite3.Row
        try:
            conexion.execute("PRAGMA busy_timeout = 10000")
            conexion.execute("PRAGMA foreign_keys = ON")
            conexion.execute("PRAGMA journal_mode = WAL")
            conexion.execute("PRAGMA synchronous = NORMAL")
            yield conexion
        finally:
            conexion.close()

    def inicializar(self):
        """Crea y migra la base existente sin eliminar registros pendientes."""
        self.directorio_app.mkdir(parents=True, exist_ok=True)
        with self._conexion() as conexion:
            conexion.execute("BEGIN IMMEDIATE")
            try:
                conexion.execute(
                    """
                    CREATE TABLE IF NOT EXISTS schema_version (
                        id INTEGER PRIMARY KEY CHECK (id = 1),
                        version INTEGER NOT NULL
                    )
                    """
                )
                conexion.execute(
                    "INSERT OR IGNORE INTO schema_version (id, version) VALUES (1, 0)"
                )
                version = conexion.execute(
                    "SELECT version FROM schema_version WHERE id = 1"
                ).fetchone()["version"]
                if version < 1:
                    self._migrar_a_v1(conexion)
                    conexion.execute(
                        "UPDATE schema_version SET version = 1 WHERE id = 1"
                    )
                    version = 1
                if version < 2:
                    self._migrar_a_v2(conexion)
                    conexion.execute(
                        "UPDATE schema_version SET version = 2 WHERE id = 1"
                    )
                    version = 2
                if version < 3:
                    self._migrar_a_v3(conexion)
                    conexion.execute(
                        "UPDATE schema_version SET version = 3 WHERE id = 1"
                    )
                    version = 3
                if version < 4:
                    self._migrar_a_v4(conexion)
                    conexion.execute(
                        "UPDATE schema_version SET version = 4 WHERE id = 1"
                    )
                    version = 4
                if version < 5:
                    self._migrar_a_v5(conexion)
                    conexion.execute(
                        "UPDATE schema_version SET version = 5 WHERE id = 1"
                    )
                    version = 5
                if version < 6:
                    self._migrar_a_v6(conexion)
                    conexion.execute(
                        "UPDATE schema_version SET version = 6 WHERE id = 1"
                    )
                    version = 6
                if version > SCHEMA_VERSION_ACTUAL:
                    raise RuntimeError(
                        "La base local fue creada por una versión más reciente"
                    )
                conexion.commit()
            except Exception:
                conexion.rollback()
                raise
        return self.ruta_db

    @staticmethod
    def _migrar_a_v1(conexion):
        """Esquema inicial. IF NOT EXISTS permite repetir una migración interrumpida."""
        conexion.execute(
            """
            CREATE TABLE IF NOT EXISTS personas_cache (
                matricula TEXT PRIMARY KEY,
                nombre TEXT NOT NULL DEFAULT '',
                rol TEXT NOT NULL DEFAULT 'ALUMNO',
                curp_ultimos_2 TEXT,
                estado TEXT,
                no_entregas INTEGER,
                grupo TEXT,
                carrera TEXT,
                fecha_actualizacion TEXT NOT NULL
            )
            """
        )

    @staticmethod
    def _migrar_a_v2(conexion):
        """Añade sesiones persistentes sin modificar eventos ya pendientes."""
        conexion.execute(
            """
            CREATE TABLE IF NOT EXISTS sesiones_locales (
                laptop_id TEXT PRIMARY KEY,
                id_sesion TEXT NOT NULL UNIQUE,
                matricula TEXT NOT NULL,
                nombre TEXT NOT NULL,
                fecha_entrada TEXT NOT NULL,
                hora_entrada TEXT NOT NULL,
                hora_salida TEXT,
                observacion TEXT NOT NULL DEFAULT 'S/N',
                origen_prestamo TEXT NOT NULL,
                prestamo_evento_id TEXT,
                devolucion_evento_id TEXT,
                no_entrega_evento_id TEXT,
                cierre_automatico_evento_id TEXT,
                fecha_creacion TEXT NOT NULL,
                fecha_actualizacion TEXT NOT NULL
            )
            """
        )

    @staticmethod
    def _migrar_a_v3(conexion):
        """Permite conservar el historial y una única sesión activa por laptop."""
        columnas = conexion.execute("PRAGMA table_info(sesiones_locales)").fetchall()
        laptop_es_pk = any(columna["name"] == "laptop_id" and columna["pk"] for columna in columnas)
        if laptop_es_pk:
            conexion.execute(
                """
                CREATE TABLE sesiones_locales_nueva (
                    id_sesion TEXT PRIMARY KEY,
                    laptop_id TEXT NOT NULL,
                    matricula TEXT NOT NULL,
                    nombre TEXT NOT NULL,
                    fecha_entrada TEXT NOT NULL,
                    hora_entrada TEXT NOT NULL,
                    hora_salida TEXT,
                    observacion TEXT NOT NULL DEFAULT 'S/N',
                    origen_prestamo TEXT NOT NULL,
                    prestamo_evento_id TEXT,
                    devolucion_evento_id TEXT,
                    no_entrega_evento_id TEXT,
                    cierre_automatico_evento_id TEXT,
                    fecha_creacion TEXT NOT NULL,
                    fecha_actualizacion TEXT NOT NULL
                )
                """
            )
            conexion.execute(
                """
                INSERT INTO sesiones_locales_nueva
                SELECT id_sesion, laptop_id, matricula, nombre, fecha_entrada,
                       hora_entrada, hora_salida, observacion, origen_prestamo,
                       prestamo_evento_id, devolucion_evento_id, no_entrega_evento_id,
                       cierre_automatico_evento_id, fecha_creacion, fecha_actualizacion
                FROM sesiones_locales
                """
            )
            conexion.execute("DROP TABLE sesiones_locales")
            conexion.execute("ALTER TABLE sesiones_locales_nueva RENAME TO sesiones_locales")
        conexion.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_sesiones_locales_una_activa
            ON sesiones_locales (laptop_id) WHERE hora_salida IS NULL
            """
        )
        conexion.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_sesiones_locales_activas
            ON sesiones_locales (matricula, hora_salida)
            """
        )

    @staticmethod
    def _migrar_a_v4(conexion):
        conexion.execute(
            """
            CREATE TABLE IF NOT EXISTS eventos_pendientes (
                id_evento TEXT PRIMARY KEY, equipo TEXT, tipo_evento TEXT NOT NULL,
                matricula TEXT, nombre TEXT, fecha_evento TEXT NOT NULL,
                hora_evento TEXT NOT NULL, laptop_id TEXT,
                datos_json TEXT NOT NULL DEFAULT '{}', sincronizado INTEGER NOT NULL DEFAULT 0,
                intentos INTEGER NOT NULL DEFAULT 0, ultimo_error TEXT,
                fecha_creacion TEXT NOT NULL
            )
            """
        )

        columnas = {
            fila["name"] for fila in conexion.execute("PRAGMA table_info(eventos_pendientes)")
        }
        if "fecha_sincronizacion" not in columnas:
            conexion.execute(
                "ALTER TABLE eventos_pendientes ADD COLUMN fecha_sincronizacion TEXT"
            )
        if "referencia_sheets" not in columnas:
            conexion.execute(
                "ALTER TABLE eventos_pendientes ADD COLUMN referencia_sheets TEXT"
            )
        conexion.execute(
            """
            CREATE TABLE IF NOT EXISTS eventos_pendientes (
                id_evento TEXT PRIMARY KEY,
                equipo TEXT,
                tipo_evento TEXT NOT NULL,
                matricula TEXT,
                nombre TEXT,
                fecha_evento TEXT NOT NULL,
                hora_evento TEXT NOT NULL,
                laptop_id TEXT,
                datos_json TEXT NOT NULL DEFAULT '{}',
                sincronizado INTEGER NOT NULL DEFAULT 0 CHECK (sincronizado IN (0, 1)),
                intentos INTEGER NOT NULL DEFAULT 0,
                ultimo_error TEXT,
                fecha_creacion TEXT NOT NULL
            )
            """
        )

        conexion.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_eventos_pendientes_sincronizado
            ON eventos_pendientes (sincronizado, fecha_creacion)
            """
        )

    @staticmethod
    def _migrar_a_v5(conexion):
        """Añade el verificador mínimo de CURP sin exponer ni borrar datos."""
        columnas = {
            fila["name"] for fila in conexion.execute("PRAGMA table_info(personas_cache)")
        }
        if "curp_ultimos_2" not in columnas:
            conexion.execute(
                "ALTER TABLE personas_cache ADD COLUMN curp_ultimos_2 TEXT"
            )

    @staticmethod
    def _migrar_a_v6(conexion):
        """Persiste el instante del fallo para diferir reintentos hist\u00f3ricos."""
        columnas = {
            fila["name"] for fila in conexion.execute("PRAGMA table_info(eventos_pendientes)")
        }
        if "fecha_ultimo_intento" not in columnas:
            conexion.execute(
                "ALTER TABLE eventos_pendientes ADD COLUMN fecha_ultimo_intento TEXT"
            )

    def version_esquema(self):
        with self._conexion() as conexion:
            fila = conexion.execute(
                "SELECT version FROM schema_version WHERE id = 1"
            ).fetchone()
        return fila["version"] if fila else None

    @staticmethod
    def _ahora():
        return datetime.now().astimezone().isoformat(timespec="microseconds")

    def upsert_persona(
        self, matricula, nombre=None, rol=None, curp_ultimos_2=None, estado=None,
        no_entregas=None, grupo=None, carrera=None, fecha_actualizacion=None,
    ):
        matricula = str(matricula).strip()
        if not matricula:
            raise ValueError("La matrícula es obligatoria para la caché local")
        fecha_actualizacion = fecha_actualizacion or self._ahora()
        with self._conexion() as conexion:
            conexion.execute(
                """
                INSERT INTO personas_cache (
                    matricula, nombre, rol, curp_ultimos_2, estado, no_entregas,
                    grupo, carrera, fecha_actualizacion
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(matricula) DO UPDATE SET
                    nombre = CASE WHEN excluded.nombre = ''
                        THEN personas_cache.nombre ELSE excluded.nombre END,
                    rol = CASE WHEN excluded.rol = ''
                        THEN personas_cache.rol ELSE excluded.rol END,
                    curp_ultimos_2 = COALESCE(excluded.curp_ultimos_2,
                                              personas_cache.curp_ultimos_2),
                    estado = COALESCE(excluded.estado, personas_cache.estado),
                    no_entregas = COALESCE(excluded.no_entregas,
                                           personas_cache.no_entregas),
                    grupo = COALESCE(excluded.grupo, personas_cache.grupo),
                    carrera = COALESCE(excluded.carrera, personas_cache.carrera),
                    fecha_actualizacion = excluded.fecha_actualizacion
                """,
                (
                    matricula, nombre or "", rol or "", curp_ultimos_2, estado,
                    no_entregas, grupo, carrera, fecha_actualizacion,
                ),
            )
        return self.obtener_persona(matricula)

    def obtener_persona(self, matricula):
        with self._conexion() as conexion:
            fila = conexion.execute(
                "SELECT * FROM personas_cache WHERE matricula = ?", (str(matricula).strip(),)
            ).fetchone()
        return dict(fila) if fila else None

    def estado_cache(self):
        with self._conexion() as conexion:
            existe = conexion.execute("SELECT 1 FROM personas_cache LIMIT 1").fetchone()
        return CACHE_OK if existe else CACHE_NO_DISPONIBLE

    def crear_evento(
        self, equipo, tipo_evento, matricula=None, nombre=None, fecha_evento=None,
        hora_evento=None, laptop_id=None, datos_json=None, id_evento=None,
        fecha_creacion=None,
    ):
        if not fecha_evento or not hora_evento:
            ahora = datetime.now().astimezone()
            fecha_evento = fecha_evento or ahora.strftime("%Y-%m-%d")
            hora_evento = hora_evento or ahora.strftime("%H:%M:%S")
        id_evento = id_evento or str(uuid.uuid4())
        if datos_json is None:
            datos_json = {}
        if not isinstance(datos_json, str):
            datos_json = json.dumps(datos_json, ensure_ascii=False, sort_keys=True)
        with self._conexion() as conexion:
            conexion.execute(
                """
                INSERT OR IGNORE INTO eventos_pendientes (
                    id_evento, equipo, tipo_evento, matricula, nombre, fecha_evento,
                    hora_evento, laptop_id, datos_json, fecha_creacion
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    id_evento, equipo, tipo_evento, matricula, nombre, fecha_evento,
                    hora_evento, laptop_id, datos_json, fecha_creacion or self._ahora(),
                ),
            )
        return id_evento

    @staticmethod
    def _fila_a_dict(fila):
        return dict(fila) if fila else None

    @staticmethod
    def _datos_fila_registro(sesion):
        return {
            "id_sesion": sesion["id_sesion"],
            "prestamo_evento_id": sesion["prestamo_evento_id"] or "",
            "matricula": sesion["matricula"],
            "nombre": sesion["nombre"],
            "fecha": sesion["fecha_entrada"],
            "hora_ingreso": sesion["hora_entrada"],
            "hora_salida": sesion["hora_salida"] or "",
            "laptop_id": sesion["laptop_id"],
            "observacion": sesion["observacion"],
        }

    @staticmethod
    def _insertar_evento(conexion, id_evento, equipo, tipo_evento, sesion, fecha_evento,
                         hora_evento, datos_json):
        conexion.execute(
            """
            INSERT OR IGNORE INTO eventos_pendientes (
                id_evento, equipo, tipo_evento, matricula, nombre, fecha_evento,
                hora_evento, laptop_id, datos_json, fecha_creacion
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                id_evento, equipo, tipo_evento, sesion["matricula"], sesion["nombre"],
                fecha_evento, hora_evento, sesion["laptop_id"],
                json.dumps(datos_json, ensure_ascii=False, sort_keys=True),
                ContingenciaLocal._ahora(),
            ),
        )

    def obtener_sesion_activa(self, laptop_id):
        with self._conexion() as conexion:
            fila = conexion.execute(
                """
                SELECT * FROM sesiones_locales
                WHERE laptop_id = ? AND hora_salida IS NULL
                """,
                (laptop_id,),
            ).fetchone()
            if not fila:
                return None
            sesion = self._fila_a_dict(fila)
            evento = sesion.get("prestamo_evento_id")
            if evento:
                sincronizado = conexion.execute(
                    "SELECT sincronizado FROM eventos_pendientes WHERE id_evento = ?", (evento,)
                ).fetchone()
                sesion["prestamo_sincronizado"] = bool(
                    sincronizado and sincronizado["sincronizado"]
                )
            else:
                sesion["prestamo_sincronizado"] = True
        return sesion

    def reconciliar_sesion_residual_local(self, id_sesion):
        """Cierra una fila residual si su DEVOLUCION ya quedó durable localmente."""
        with self._conexion() as conexion:
            conexion.execute("BEGIN IMMEDIATE")
            try:
                sesion = conexion.execute(
                    "SELECT * FROM sesiones_locales WHERE id_sesion = ?", (id_sesion,)
                ).fetchone()
                if not sesion or sesion["hora_salida"] or not sesion["devolucion_evento_id"]:
                    conexion.commit()
                    return False
                evento = conexion.execute(
                    """
                    SELECT hora_evento FROM eventos_pendientes
                    WHERE id_evento = ? AND tipo_evento = 'DEVOLUCION'
                    """,
                    (sesion["devolucion_evento_id"],),
                ).fetchone()
                if not evento:
                    conexion.commit()
                    return False
                conexion.execute(
                    """
                    UPDATE sesiones_locales
                    SET hora_salida = ?, fecha_actualizacion = ?
                    WHERE id_sesion = ? AND hora_salida IS NULL
                    """,
                    (evento["hora_evento"], self._ahora(), id_sesion),
                )
                conexion.commit()
                return True
            except Exception:
                conexion.rollback()
                raise

    def cerrar_sesion_residual_remota(self, id_sesion, hora_salida):
        """Refleja un cierre ya confirmado en Sheets sin crear otro evento."""
        with self._conexion() as conexion:
            cursor = conexion.execute(
                """
                UPDATE sesiones_locales
                SET hora_salida = ?, fecha_actualizacion = ?
                WHERE id_sesion = ? AND hora_salida IS NULL
                """,
                (hora_salida, self._ahora(), id_sesion),
            )
        return bool(cursor.rowcount)

    def registrar_prestamo_online(self, matricula, nombre, laptop_id, fecha_entrada,
                                  hora_entrada):
        """Persiste la sesión confirmada por Sheets para una devolución posterior offline."""
        return self._registrar_prestamo(
            matricula, nombre, laptop_id, fecha_entrada, hora_entrada, "ONLINE", None
        )

    def registrar_prestamo_offline(self, matricula, nombre, laptop_id, fecha_entrada,
                                   hora_entrada, equipo=None):
        return self._registrar_prestamo(
            matricula, nombre, laptop_id, fecha_entrada, hora_entrada,
            "OFFLINE", equipo or laptop_id,
        )

    def registrar_prestamo_coordinador(self, matricula, nombre, laptop_id, fecha_entrada,
                                       hora_entrada, equipo=None):
        """Crea el UUID antes del POST online; queda pendiente si éste falla."""
        return self._registrar_prestamo(
            matricula, nombre, laptop_id, fecha_entrada, hora_entrada,
            "COORDINADOR", equipo or laptop_id,
        )

    def _registrar_prestamo(self, matricula, nombre, laptop_id, fecha_entrada,
                            hora_entrada, origen, equipo):
        """Inserta sesión y PRESTAMO en una sola transacción, idempotente por laptop."""
        with self._conexion() as conexion:
            conexion.execute("BEGIN IMMEDIATE")
            try:
                existente, sesion_reconciliada = self._reconciliar_sesion_activa_en_conexion(
                    conexion, laptop_id
                )
                if existente:
                    if existente["matricula"] != str(matricula).strip():
                        raise RuntimeError("Ya existe una sesión activa en esta laptop")
                    resultado = self._fila_a_dict(existente)
                    resultado["sesion_existente"] = True
                    resultado["sesion_reconciliada"] = sesion_reconciliada
                    conexion.commit()
                    return resultado
                ahora = self._ahora()
                id_sesion = str(uuid.uuid4())
                id_evento = str(uuid.uuid4()) if origen in ("OFFLINE", "COORDINADOR") else None
                conexion.execute(
                    """
                    INSERT INTO sesiones_locales (
                        laptop_id, id_sesion, matricula, nombre, fecha_entrada,
                        hora_entrada, origen_prestamo, prestamo_evento_id,
                        fecha_creacion, fecha_actualizacion
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        laptop_id, id_sesion, str(matricula).strip(), nombre,
                        fecha_entrada, hora_entrada, origen, id_evento, ahora, ahora,
                    ),
                )
                if id_sesion:
                    sesion = conexion.execute(
                        "SELECT * FROM sesiones_locales WHERE id_sesion = ?", (id_sesion,)
                    ).fetchone()
                else:
                    sesion = conexion.execute(
                        """
                        SELECT * FROM sesiones_locales
                        WHERE laptop_id = ? AND matricula = ?
                        ORDER BY fecha_creacion DESC, id_sesion DESC LIMIT 1
                        """,
                        (laptop_id, str(matricula).strip()),
                    ).fetchone()
                if id_evento:
                    self._insertar_evento(
                        conexion, id_evento, equipo, "PRESTAMO", sesion,
                        fecha_entrada, hora_entrada, self._datos_fila_registro(sesion),
                    )
                conexion.commit()
                resultado = self._fila_a_dict(sesion)
                resultado["sesion_existente"] = False
                resultado["sesion_reconciliada"] = sesion_reconciliada
                return resultado
            except Exception:
                conexion.rollback()
                raise

    @staticmethod
    def _reconciliar_sesion_activa_en_conexion(conexion, laptop_id):
        """Cierra sólo una sesión cuyo cierre ya fue aplicado remotamente.

        Una sesión sin evidencia de devolución/no-entrega/cierre sincronizado se
        conserva: podría corresponder a un préstamo real todavía abierto.
        """
        sesion = conexion.execute(
            "SELECT * FROM sesiones_locales WHERE laptop_id = ? AND hora_salida IS NULL",
            (laptop_id,),
        ).fetchone()
        if not sesion:
            return None, False
        cierres = (
            ("devolucion_evento_id", None),
            ("no_entrega_evento_id", "NO_ENTREGA_CONTADA"),
            ("cierre_automatico_evento_id", "CIERRE_AUTOMATICO_POR_NUEVA_SESION"),
        )
        for columna, observacion in cierres:
            id_evento = sesion[columna]
            if not id_evento:
                continue
            evento = conexion.execute(
                """
                SELECT tipo_evento, hora_evento, sincronizado FROM eventos_pendientes
                WHERE id_evento = ?
                """,
                (id_evento,),
            ).fetchone()
            if not evento or not evento["sincronizado"]:
                continue
            actualizaciones = "hora_salida = ?, fecha_actualizacion = ?"
            parametros = [evento["hora_evento"], ContingenciaLocal._ahora()]
            if observacion:
                actualizaciones += ", observacion = ?"
                parametros.append(observacion)
            parametros.append(sesion["id_sesion"])
            conexion.execute(
                f"UPDATE sesiones_locales SET {actualizaciones} WHERE id_sesion = ?",
                parametros,
            )
            return None, True
        return sesion, False

    def registrar_devolucion_offline(self, matricula, laptop_id, fecha_evento,
                                    hora_evento, equipo=None, id_sesion=None):
        """Cierra una sesión local y genera una sola DEVOLUCION pendiente."""
        with self._conexion() as conexion:
            conexion.execute("BEGIN IMMEDIATE")
            try:
                if id_sesion:
                    sesion = conexion.execute(
                        "SELECT * FROM sesiones_locales WHERE id_sesion = ?", (id_sesion,)
                    ).fetchone()
                else:
                    sesion = conexion.execute(
                        """
                        SELECT * FROM sesiones_locales
                        WHERE laptop_id = ? AND matricula = ?
                        ORDER BY fecha_creacion DESC, id_sesion DESC LIMIT 1
                        """,
                        (laptop_id, str(matricula).strip()),
                    ).fetchone()
                if (
                    not sesion
                    or sesion["matricula"] != str(matricula).strip()
                    or sesion["laptop_id"] != laptop_id
                ):
                    raise RuntimeError("No existe una sesión local activa para esta laptop")
                if sesion["hora_salida"]:
                    if sesion["devolucion_evento_id"]:
                        conexion.commit()
                        return self._fila_a_dict(sesion)
                    raise RuntimeError("La sesión local ya está cerrada")
                id_evento = sesion["devolucion_evento_id"] or str(uuid.uuid4())
                conexion.execute(
                    """
                    UPDATE sesiones_locales
                    SET hora_salida = ?, devolucion_evento_id = ?, fecha_actualizacion = ?
                    WHERE id_sesion = ?
                    """,
                    (hora_evento, id_evento, self._ahora(), sesion["id_sesion"]),
                )
                sesion = conexion.execute(
                    "SELECT * FROM sesiones_locales WHERE id_sesion = ?", (sesion["id_sesion"],)
                ).fetchone()
                self._insertar_evento(
                    conexion, id_evento, equipo or laptop_id, "DEVOLUCION", sesion,
                    fecha_evento, hora_evento, self._datos_fila_registro(sesion),
                )
                conexion.commit()
                return self._fila_a_dict(sesion)
            except Exception:
                conexion.rollback()
                raise

    def registrar_devolucion_online(self, matricula, laptop_id, hora_salida,
                                    id_sesion=None):
        """Refleja una devolución ya confirmada por Sheets sin crear evento local."""
        with self._conexion() as conexion:
            if id_sesion:
                conexion.execute(
                    """
                    UPDATE sesiones_locales
                    SET hora_salida = COALESCE(hora_salida, ?), fecha_actualizacion = ?
                    WHERE id_sesion = ? AND laptop_id = ? AND matricula = ?
                      AND hora_salida IS NULL
                    """,
                    (hora_salida, self._ahora(), id_sesion, laptop_id, str(matricula).strip()),
                )
            else:
                conexion.execute(
                    """
                    UPDATE sesiones_locales
                    SET hora_salida = COALESCE(hora_salida, ?), fecha_actualizacion = ?
                    WHERE laptop_id = ? AND matricula = ? AND hora_salida IS NULL
                    """,
                    (hora_salida, self._ahora(), laptop_id, str(matricula).strip()),
                )
        return self.obtener_sesion_activa(laptop_id)

    def registrar_no_entrega_local(self, matricula, laptop_id, fecha_evento,
                                   hora_evento, equipo=None, id_sesion=None):
        """Genera un NO_ENTREGA a lo sumo una vez por sesión local."""
        return self._cerrar_sesion_local(
            matricula, laptop_id, fecha_evento, hora_evento, equipo,
            "no_entrega_evento_id", "NO_ENTREGA", "NO_ENTREGA_CONTADA", id_sesion,
        )

    def registrar_cierre_automatico_local(self, matricula, laptop_id, fecha_evento,
                                          hora_evento, equipo=None, id_sesion=None):
        return self._cerrar_sesion_local(
            matricula, laptop_id, fecha_evento, hora_evento, equipo,
            "cierre_automatico_evento_id", "CIERRE_AUTOMATICO",
            "CIERRE_AUTOMATICO_POR_NUEVA_SESION", id_sesion,
        )

    def cerrar_sesion_residual_al_arrancar(self, laptop_id, fecha_evento,
                                            hora_evento, equipo=None):
        """Cierra atómicamente una sesión heredada con sus dos eventos.

        No depende de Sheets. Los UUID deterministas y las columnas de sesión
        hacen que reinicios repetidos no creen NO_ENTREGA duplicados.
        """
        with self._conexion() as conexion:
            conexion.execute("BEGIN IMMEDIATE")
            try:
                sesion = conexion.execute(
                    """
                    SELECT * FROM sesiones_locales
                    WHERE laptop_id = ? AND hora_salida IS NULL
                    """,
                    (laptop_id,),
                ).fetchone()
                if not sesion:
                    conexion.commit()
                    return {"resultado": "SIN_SESION", "sesion": None}

                id_no_entrega = sesion["no_entrega_evento_id"] or str(uuid.uuid5(
                    uuid.NAMESPACE_URL, f"utp-v141:NO_ENTREGA:{sesion['id_sesion']}"
                ))
                id_cierre = sesion["cierre_automatico_evento_id"] or str(uuid.uuid5(
                    uuid.NAMESPACE_URL,
                    f"utp-v141:CIERRE_AUTOMATICO:{sesion['id_sesion']}",
                ))
                crear_no_entrega = not sesion["no_entrega_evento_id"]
                crear_cierre = not sesion["cierre_automatico_evento_id"]
                conexion.execute(
                    """
                    UPDATE sesiones_locales
                    SET hora_salida = ?, observacion = ?, no_entrega_evento_id = ?,
                        cierre_automatico_evento_id = ?, fecha_actualizacion = ?
                    WHERE id_sesion = ?
                    """,
                    (
                        hora_evento, "CIERRE_AUTOMATICO_POR_REINICIO",
                        id_no_entrega, id_cierre, self._ahora(), sesion["id_sesion"],
                    ),
                )
                sesion_cerrada = conexion.execute(
                    "SELECT * FROM sesiones_locales WHERE id_sesion = ?",
                    (sesion["id_sesion"],),
                ).fetchone()
                if crear_no_entrega:
                    self._insertar_evento(
                        conexion, id_no_entrega, equipo or laptop_id, "NO_ENTREGA",
                        sesion_cerrada, fecha_evento, hora_evento,
                        self._datos_fila_registro(sesion_cerrada),
                    )
                if crear_cierre:
                    self._insertar_evento(
                        conexion, id_cierre, equipo or laptop_id, "CIERRE_AUTOMATICO",
                        sesion_cerrada, fecha_evento, hora_evento,
                        self._datos_fila_registro(sesion_cerrada),
                    )
                conexion.commit()
                return {
                    "resultado": "SESION_RESIDUAL_CERRADA",
                    "sesion": self._fila_a_dict(sesion_cerrada),
                    "no_entrega_creada": crear_no_entrega,
                    "cierre_creado": crear_cierre,
                }
            except Exception:
                conexion.rollback()
                raise

    def resolver_sesion_activa_por_nueva_matricula(
        self, laptop_id, matricula, fecha_evento, hora_evento, equipo=None,
        registrar_no_entrega=True,
    ):
        """Conserva la misma sesión o cierra la anterior para otro alumno.

        La resolución ocurre en SQLite antes de que el llamador envíe los UUID
        persistidos al coordinador; repetirla no genera eventos duplicados.
        """
        sesion = self.obtener_sesion_activa(laptop_id)
        if not sesion:
            return {"resultado": "SIN_SESION", "sesion": None}
        if sesion["matricula"] == str(matricula).strip():
            return {"resultado": "MISMA_MATRICULA", "sesion": sesion}

        matricula_anterior = sesion["matricula"]
        if registrar_no_entrega:
            self.registrar_no_entrega_local(
                matricula_anterior, laptop_id, fecha_evento, hora_evento, equipo,
                id_sesion=sesion["id_sesion"],
            )
        cerrada = self.registrar_cierre_automatico_local(
            matricula_anterior, laptop_id, fecha_evento, hora_evento, equipo,
            id_sesion=sesion["id_sesion"],
        )
        return {"resultado": "SESION_ANTERIOR_CERRADA", "sesion": cerrada}

    def _cerrar_sesion_local(self, matricula, laptop_id, fecha_evento, hora_evento,
                              equipo, columna_evento, tipo_evento, observacion,
                              id_sesion=None):
        with self._conexion() as conexion:
            conexion.execute("BEGIN IMMEDIATE")
            try:
                if id_sesion:
                    sesion = conexion.execute(
                        "SELECT * FROM sesiones_locales WHERE id_sesion = ?", (id_sesion,)
                    ).fetchone()
                else:
                    sesion = conexion.execute(
                        """
                        SELECT * FROM sesiones_locales
                        WHERE laptop_id = ? AND matricula = ?
                        ORDER BY fecha_creacion DESC, id_sesion DESC LIMIT 1
                        """,
                        (laptop_id, str(matricula).strip()),
                    ).fetchone()
                if (
                    not sesion
                    or sesion["matricula"] != str(matricula).strip()
                    or sesion["laptop_id"] != laptop_id
                ):
                    raise RuntimeError("No existe una sesión local para cerrar")
                id_evento = sesion[columna_evento]
                if id_evento:
                    conexion.commit()
                    return self._fila_a_dict(sesion)
                id_evento = str(uuid.uuid5(
                    uuid.NAMESPACE_URL, f"utp-v141:{tipo_evento}:{sesion['id_sesion']}"
                ))
                conexion.execute(
                    f"""
                    UPDATE sesiones_locales
                    SET hora_salida = COALESCE(hora_salida, ?), observacion = ?,
                        {columna_evento} = ?, fecha_actualizacion = ?
                    WHERE id_sesion = ?
                    """,
                    (hora_evento, observacion, id_evento, self._ahora(), sesion["id_sesion"]),
                )
                sesion = conexion.execute(
                    "SELECT * FROM sesiones_locales WHERE id_sesion = ?", (sesion["id_sesion"],)
                ).fetchone()
                self._insertar_evento(
                    conexion, id_evento, equipo or laptop_id, tipo_evento, sesion,
                    fecha_evento, hora_evento, self._datos_fila_registro(sesion),
                )
                conexion.commit()
                return self._fila_a_dict(sesion)
            except Exception:
                conexion.rollback()
                raise

    def no_entregas_efectivas(self, matricula, no_entregas_cache):
        with self._conexion() as conexion:
            fila = conexion.execute(
                """
                SELECT COUNT(*) AS total FROM eventos_pendientes
                WHERE matricula = ? AND tipo_evento = 'NO_ENTREGA' AND sincronizado = 0
                """,
                (str(matricula).strip(),),
            ).fetchone()
        return int(no_entregas_cache or 0) + int(fila["total"])

    def obtener_evento(self, id_evento):
        with self._conexion() as conexion:
            fila = conexion.execute(
                "SELECT * FROM eventos_pendientes WHERE id_evento = ?", (id_evento,)
            ).fetchone()
        return dict(fila) if fila else None

    def obtener_eventos_pendientes(self):
        with self._conexion() as conexion:
            filas = conexion.execute(
                """
                SELECT * FROM eventos_pendientes
                WHERE sincronizado = 0
                ORDER BY CASE WHEN intentos = 0 THEN 0 ELSE 1 END,
                         fecha_creacion, id_evento
                """
            ).fetchall()
        return [dict(fila) for fila in filas]

    def incrementar_intentos(self, id_evento, ultimo_error=None):
        with self._conexion() as conexion:
            conexion.execute(
                """
                UPDATE eventos_pendientes
                SET intentos = intentos + 1,
                    ultimo_error = COALESCE(?, ultimo_error),
                    fecha_ultimo_intento = ?
                WHERE id_evento = ?
                """,
                (ultimo_error, self._ahora(), id_evento),
            )
        return self.obtener_evento(id_evento)

    def guardar_ultimo_error(self, id_evento, ultimo_error):
        with self._conexion() as conexion:
            conexion.execute(
                "UPDATE eventos_pendientes SET ultimo_error = ? WHERE id_evento = ?",
                (ultimo_error, id_evento),
            )
        return self.obtener_evento(id_evento)

    def marcar_sincronizado(self, id_evento, referencia_sheets=None):
        with self._conexion() as conexion:
            conexion.execute(
                """
                UPDATE eventos_pendientes
                SET sincronizado = 1, fecha_sincronizacion = ?,
                    referencia_sheets = COALESCE(?, referencia_sheets)
                WHERE id_evento = ?
                """,
                (self._ahora(), referencia_sheets, id_evento),
            )
        return self.obtener_evento(id_evento)
