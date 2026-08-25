import tkinter as tk
import threading
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import socket
import requests
import pytz
import os
import time
import traceback
import hashlib
import hmac
import subprocess
import sys
import certifi
import tempfile
import msvcrt
from contextlib import contextmanager
from proximidad_carrito import (
    ESTADO_ERROR,
    ESTADO_NO_DETECTADO,
    verificar_proximidad_carrito
)
from heartbeat_manager import iniciar_heartbeat_manager

DIRECTORIO_APP = os.path.dirname(os.path.abspath(__file__))
RUTA_CREDENCIALES = os.path.join(DIRECTORIO_APP, "credenciales.json")
RUTA_UTP = os.path.join(DIRECTORIO_APP, "UTP.png")

# Intentar importar PIL, si no está disponible usar emojis
try:
    from PIL import Image, ImageTk, ImageOps, ImageChops
    PIL_DISPONIBLE = True
except ImportError:
    PIL_DISPONIBLE = False

# --- VARIABLES GLOBALES ---
VERSION_SISTEMA = "1.3.3"
MODO_PRUEBA = False
PROXIMIDAD_HABILITADA = False
hoja_alumnos = None
hoja_registros = None
zona_horaria = pytz.timezone("America/Chihuahua")
verificar_conexion_activo = True
procesando_sesion = False
aviso_internet = None
acepta_estado_equipo = None
chk_label = None
verificacion_aviso_en_curso = False
conexion_google_lock = threading.Lock()
actualizacion_lock = threading.Lock()
dialogo_admin_abierto = False
contexto_contingencia_proximidad = None
confirmacion_contingencia_abierta = False
registro_contingencias_sesion = []
ventana_entrega_activa = None

TIEMPO_INACTIVIDAD_MS = 60000
timeout_inactividad_id = None
timeout_inactividad_activo = False
flujo_cancelado_por_inactividad = False
ventanas_flujo_alumno = set()

HASH_CLAVE_ADMIN = "0da417ade5c1bd3c262b600330f15e11af9b87a5b3984caac3d136a307b63190"
ARCHIVO_LOG_CONTINGENCIA = os.path.join(
    DIRECTORIO_APP, "contingencia_admin.log"
)
ARCHIVO_LOG_ACTUALIZACION = "update.log"
ARCHIVO_LOG_INICIO = "inicio_tecnico.log"
cierre_aplicacion_en_curso = False


class ErrorConsultaGoogle(Exception):
    """Indica que Google Sheets no pudo entregar un resultado confiable."""


def registrar_evento_tecnico(evento):
    directorio = os.path.dirname(os.path.abspath(__file__))
    ruta_log = os.path.join(directorio, ARCHIVO_LOG_INICIO)
    try:
        with open(ruta_log, "a", encoding="utf-8") as archivo:
            archivo.write(
                f"[{datetime.now().astimezone().isoformat(timespec='seconds')}] "
                f"evento={evento}\tpid={os.getpid()}\thostname={socket.gethostname()}\n"
            )
    except OSError:
        pass


def registrar_inicio_tecnico():
    directorio = os.path.dirname(os.path.abspath(__file__))
    commit = "NO_DISPONIBLE"
    try:
        resultado = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=directorio,
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=3,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            check=False
        )
        if resultado.returncode == 0:
            commit = resultado.stdout.strip() or commit
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass

    ruta_log = os.path.join(directorio, ARCHIVO_LOG_INICIO)
    try:
        with open(ruta_log, "a", encoding="utf-8") as archivo:
            archivo.write(
                f"[{datetime.now().astimezone().isoformat(timespec='seconds')}] "
                f"ruta={os.path.abspath(__file__)}\tpid={os.getpid()}\t"
                f"hostname={socket.gethostname()}\tcommit={commit}\n"
            )
    except OSError:
        pass

# =========================
# COLUMNAS HOJA REGISTROS (0-based)
# =========================
COL_MATRICULA = 0
COL_NOMBRE = 1
COL_FECHA = 2
COL_HORA_INGRESO = 3
COL_HORA_SALIDA = 4
COL_LAPTOP_ID = 5
COL_OBSERVACION = 6


def observacion_es_normal(valor):
    return (valor or "").strip() in ("", "S/N")


ESTADO_VACIO_VALIDO = "ESTADO_VACIO_VALIDO"


def normalizar_estado_alumno(valor, lectura_exitosa=False):
    if valor == ESTADO_VACIO_VALIDO:
        return ESTADO_VACIO_VALIDO
    if valor is None:
        return ESTADO_VACIO_VALIDO if lectura_exitosa else None

    estado = str(valor).strip().upper()
    if estado == "":
        return ESTADO_VACIO_VALIDO
    return estado if estado in ("ACTIVO", "ADVERTENCIA", "SANCIONADO") else None


def estado_alumno_permite_prestamo(valor):
    return normalizar_estado_alumno(valor) in {
        ESTADO_VACIO_VALIDO,
        "ACTIVO",
        "ADVERTENCIA",
    }


@contextmanager
def bloqueo_local_registro_prestamo():
    """Serializa el tramo leer/append entre procesos de esta laptop."""
    identificador = hashlib.sha256(
        os.path.abspath(__file__).encode("utf-8")
    ).hexdigest()[:16]
    ruta = os.path.join(tempfile.gettempdir(), f"registro_laptop_{identificador}.lock")
    with open(ruta, "a+b") as archivo:
        archivo.seek(0, os.SEEK_END)
        if archivo.tell() == 0:
            archivo.write(b"0")
            archivo.flush()
        archivo.seek(0)
        msvcrt.locking(archivo.fileno(), msvcrt.LK_LOCK, 1)
        try:
            yield
        finally:
            archivo.seek(0)
            msvcrt.locking(archivo.fileno(), msvcrt.LK_UNLCK, 1)

MATRICULAS_EXENTAS_NO_ENTREGA = {
    "010537700",
    "082025",
    "11",
}
# Cambiar a "clasico" permite recuperar la apariencia anterior sin tocar código.
TEMA_INTERFAZ = "moderno"

TEMAS_INTERFAZ = {
    "clasico": {
        "colores": {
            "primario": "#0066cc", "secundario": "#00a8ff",
            "fondo": "#f5f7fa", "tarjeta": "#ffffff",
            "texto": "#2d3436", "texto_secundario": "#636e72",
            "exito": "#00b894", "error": "#e84118",
            "advertencia": "#fdcb6e", "borde": "#dfe6e9",
            "hover": "#f1f2f6", "campo": "#f8f9fa",
            "info_suave": "#eef6ff", "advertencia_suave": "#fff4d6"
        },
        "fuentes": {
            "titulo": ("Segoe UI", 24, "bold"),
            "subtitulo": ("Segoe UI", 14, "bold"),
            "cuerpo": ("Segoe UI", 11),
            "cuerpo_bold": ("Segoe UI", 11, "bold"),
            "entrada": ("Segoe UI", 14),
            "boton": ("Segoe UI", 12, "bold"),
            "pequena": ("Segoe UI", 9),
            "pequena_bold": ("Segoe UI", 9, "bold")
        },
        "medidas": {"tarjeta_login_ancho": 460, "tarjeta_login_alto": 620,
                    "dialogo_ancho": 440, "radio_borde": 1},
        "marca_agua_fondo_login": False
    },
    "moderno": {
        "colores": {
            "primario": "#124E78", "secundario": "#0E7490",
            "fondo": "#EDF2F7", "tarjeta": "#FFFFFF",
            "texto": "#172B3A", "texto_secundario": "#5F7180",
            "exito": "#21835B", "error": "#C2413B",
            "advertencia": "#D89B20", "borde": "#D7E0E8",
            "hover": "#E7EEF4", "campo": "#F7F9FB",
            "info_suave": "#EAF3F8", "advertencia_suave": "#FFF6DE"
        },
        "fuentes": {
            "titulo": ("Segoe UI", 26, "bold"),
            "subtitulo": ("Segoe UI", 15, "bold"),
            "cuerpo": ("Segoe UI", 11),
            "cuerpo_bold": ("Segoe UI", 11, "bold"),
            "entrada": ("Segoe UI", 15),
            "boton": ("Segoe UI", 11, "bold"),
            "pequena": ("Segoe UI", 9),
            "pequena_bold": ("Segoe UI", 9, "bold")
        },
        "medidas": {"tarjeta_login_ancho": 480, "tarjeta_login_alto": 640,
                    "dialogo_ancho": 460, "radio_borde": 1},
        "marca_agua_fondo_login": True
    }
}

if TEMA_INTERFAZ not in TEMAS_INTERFAZ:
    TEMA_INTERFAZ = "clasico"

ESTILO_ACTUAL = TEMAS_INTERFAZ[TEMA_INTERFAZ]
COLORES = ESTILO_ACTUAL["colores"]
FUENTES = ESTILO_ACTUAL["fuentes"]
TAMANOS = ESTILO_ACTUAL["medidas"]
FONDO_MARCA_AGUA_LOGIN = ESTILO_ACTUAL["marca_agua_fondo_login"]

COLOR_PRIMARIO = COLORES["primario"]
COLOR_SECUNDARIO = COLORES["secundario"]
COLOR_FONDO = COLORES["fondo"]
COLOR_TARJETA = COLORES["tarjeta"]
COLOR_TEXTO = COLORES["texto"]
COLOR_TEXTO_SECUNDARIO = COLORES["texto_secundario"]
COLOR_EXITO = COLORES["exito"]
COLOR_ERROR = COLORES["error"]
COLOR_ADVERTENCIA = COLORES["advertencia"]
COLOR_BORDE = COLORES["borde"]
COLOR_HOVER = COLORES["hover"]

FUENTE_TITULO = FUENTES["titulo"]
FUENTE_SUBTITULO = FUENTES["subtitulo"]
FUENTE_CUERPO = FUENTES["cuerpo"]
FUENTE_CUERPO_BOLD = FUENTES["cuerpo_bold"]
FUENTE_ENTRADA = FUENTES["entrada"]
FUENTE_BOTON = FUENTES["boton"]
FUENTE_PEQ = FUENTES["pequena"]
FUENTE_PEQ_BOLD = FUENTES["pequena_bold"]


def estilo_boton(tipo="principal"):
    fondos = {
        "principal": COLOR_PRIMARIO,
        "secundario": COLOR_HOVER,
        "exito": COLOR_EXITO,
        "error": COLOR_ERROR
    }
    fondos_activos = {
        "principal": COLOR_SECUNDARIO,
        "secundario": COLOR_BORDE,
        "exito": "#176B49",
        "error": "#9F332F"
    }
    fondo = fondos.get(tipo, COLOR_PRIMARIO)
    return {
        "font": FUENTE_BOTON,
        "bg": fondo,
        "fg": COLOR_TEXTO if tipo == "secundario" else "white",
        "activebackground": fondos_activos.get(tipo, COLOR_SECUNDARIO),
        "activeforeground": COLOR_TEXTO if tipo == "secundario" else "white",
        "bd": 0,
        "relief": tk.FLAT,
        "cursor": "hand2"
    }


def estilo_entrada():
    return {
        "font": FUENTE_ENTRADA,
        "bg": COLORES["campo"],
        "fg": COLOR_TEXTO,
        "insertbackground": COLOR_TEXTO,
        "relief": tk.FLAT,
        "highlightthickness": 2,
        "highlightbackground": COLOR_BORDE,
        "highlightcolor": COLOR_PRIMARIO
    }


def _centrar_dialogo(
    dialogo,
    parent,
    ancho_minimo=440,
    alto_minimo=190,
    margen_horizontal=12,
    margen_vertical=12
):
    dialogo.update_idletasks()
    ancho_requerido = dialogo.winfo_reqwidth()
    alto_requerido = dialogo.winfo_reqheight()
    ancho = max(ancho_requerido + margen_horizontal, ancho_minimo)
    alto = max(alto_requerido + margen_vertical, alto_minimo)

    parent.update_idletasks()
    ancho_parent = parent.winfo_width()
    alto_parent = parent.winfo_height()

    if ancho_parent > 1 and alto_parent > 1:
        x = parent.winfo_rootx() + (ancho_parent - ancho) // 2
        y = parent.winfo_rooty() + (alto_parent - alto) // 2
    else:
        x = (dialogo.winfo_screenwidth() - ancho) // 2
        y = (dialogo.winfo_screenheight() - alto) // 2

    dialogo.geometry(f"{ancho}x{alto}+{x}+{y}")
    dialogo.deiconify()
    dialogo.lift()


def _crear_dialogo_personalizado(titulo, mensaje, tipo, parent, texto_destacado=None):
    estilos = {
        "info": (COLOR_PRIMARIO, "i"),
        "advertencia": ("#f0ad00", "!"),
        "error": (COLOR_ERROR, "×"),
        "exito": (COLOR_EXITO, "✓")
    }
    color, icono = estilos.get(tipo, estilos["info"])
    parent = parent or ventana

    dialogo = tk.Toplevel(parent)
    dialogo.title(titulo)
    dialogo.resizable(False, False)
    dialogo.configure(bg=COLOR_FONDO)
    dialogo.transient(parent)
    dialogo.attributes("-topmost", True)
    dialogo.withdraw()

    tarjeta = tk.Frame(
        dialogo,
        bg=COLOR_TARJETA,
        highlightthickness=1,
        highlightbackground=COLOR_BORDE
    )
    tarjeta.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    if TEMA_INTERFAZ == "moderno":
        tk.Frame(tarjeta, height=4, bg=color).pack(fill=tk.X, side=tk.TOP)

    botones = tk.Frame(tarjeta, bg=COLOR_TARJETA, padx=22)
    botones.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 18))

    cuerpo = tk.Frame(tarjeta, bg=COLOR_TARJETA, padx=22, pady=20)
    cuerpo.pack(fill=tk.BOTH, expand=True)

    icono_canvas = tk.Canvas(
        cuerpo,
        width=56,
        height=56,
        bg=COLOR_TARJETA,
        highlightthickness=0
    )
    icono_canvas.pack(side=tk.LEFT, anchor="n", padx=(0, 18))
    icono_canvas.create_oval(4, 4, 52, 52, fill=color, outline=color)
    icono_canvas.create_text(
        28,
        28,
        text=icono,
        fill="white",
        font=("Segoe UI", 21, "bold")
    )

    texto = tk.Frame(cuerpo, bg=COLOR_TARJETA)
    texto.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    tk.Label(
        texto,
        text=titulo,
        font=("Segoe UI", 15, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA,
        anchor="w"
    ).pack(fill=tk.X, pady=(2, 8))
    tk.Label(
        texto,
        text=mensaje,
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA,
        justify=tk.LEFT,
        anchor="w",
        wraplength=360
    ).pack(fill=tk.X)

    if texto_destacado:
        tk.Label(
            texto,
            text=texto_destacado,
            font=FUENTE_CUERPO_BOLD,
            fg="#8A5A00",
            bg=COLORES["advertencia_suave"],
            justify=tk.LEFT,
            anchor="w",
            wraplength=340,
            padx=12,
            pady=9
        ).pack(fill=tk.X, pady=(12, 0))

    return dialogo, botones, color, parent


def mostrar_dialogo_personalizado(
    titulo,
    mensaje,
    tipo="info",
    parent=None,
    al_crear=None,
    texto_destacado=None,
    texto_boton="ACEPTAR"
):
    timeout_iniciado_por_dialogo = False
    if procesando_sesion and not timeout_inactividad_activo:
        activar_timeout_inactividad()
        timeout_iniciado_por_dialogo = True

    dialogo, botones, color, parent = _crear_dialogo_personalizado(
        titulo, mensaje, tipo, parent, texto_destacado
    )
    registrar_ventana_flujo_alumno(dialogo)
    grab_anterior = parent.grab_current()

    def cerrar():
        try:
            dialogo.grab_release()
        except tk.TclError:
            pass
        dialogo.destroy()
        try:
            if grab_anterior and grab_anterior.winfo_exists():
                grab_anterior.grab_set()
        except tk.TclError:
            pass

    tk.Button(
        botones,
        text=texto_boton,
        font=FUENTE_BOTON,
        bg=color,
        fg="white",
        activebackground=color,
        activeforeground="white",
        bd=0,
        cursor="hand2",
        padx=26,
        pady=8,
        command=cerrar
    ).pack(side=tk.RIGHT)

    dialogo.protocol("WM_DELETE_WINDOW", cerrar)
    dialogo.bind("<Return>", lambda _evento: cerrar())
    dialogo.bind("<Escape>", lambda _evento: cerrar())
    if al_crear:
        al_crear(dialogo, cerrar)
    _centrar_dialogo(dialogo, parent, ancho_minimo=TAMANOS["dialogo_ancho"])
    dialogo.grab_set()
    dialogo.focus_force()
    dialogo.wait_window()
    if timeout_iniciado_por_dialogo:
        cancelar_timeout_inactividad()


def mostrar_confirmacion_personalizada(titulo, mensaje, parent=None):
    dialogo, botones, color, parent = _crear_dialogo_personalizado(
        titulo, mensaje, "info", parent
    )
    registrar_ventana_flujo_alumno(dialogo)
    grab_anterior = parent.grab_current()
    resultado = {"valor": False}

    def cerrar(valor=False):
        resultado["valor"] = valor
        try:
            dialogo.grab_release()
        except tk.TclError:
            pass
        dialogo.destroy()
        try:
            if grab_anterior and grab_anterior.winfo_exists():
                grab_anterior.grab_set()
        except tk.TclError:
            pass

    tk.Button(
        botones,
        text="NO",
        padx=26,
        pady=8,
        command=lambda: cerrar(False),
        **estilo_boton("secundario")
    ).pack(side=tk.RIGHT, padx=(12, 0))
    tk.Button(
        botones,
        text="SÍ",
        padx=26,
        pady=8,
        command=lambda: cerrar(True),
        **estilo_boton("principal")
    ).pack(side=tk.RIGHT)

    dialogo.protocol("WM_DELETE_WINDOW", lambda: cerrar(False))
    dialogo.bind("<Return>", lambda _evento: cerrar(True))
    dialogo.bind("<Escape>", lambda _evento: cerrar(False))
    _centrar_dialogo(dialogo, parent, ancho_minimo=TAMANOS["dialogo_ancho"])
    dialogo.grab_set()
    dialogo.focus_force()
    dialogo.wait_window()
    return resultado["valor"]


def mostrar_informacion(titulo, mensaje, parent=None):
    return mostrar_dialogo_personalizado(titulo, mensaje, "info", parent)


def mostrar_advertencia(titulo, mensaje, parent=None):
    return mostrar_dialogo_personalizado(titulo, mensaje, "advertencia", parent)


def mostrar_error(
    titulo,
    mensaje,
    parent=None,
    al_crear=None,
    texto_destacado=None,
    texto_boton="ACEPTAR"
):
    return mostrar_dialogo_personalizado(
        titulo,
        mensaje,
        "error",
        parent,
        al_crear,
        texto_destacado,
        texto_boton
    )


def mostrar_exito(titulo, mensaje, parent=None):
    return mostrar_dialogo_personalizado(titulo, mensaje, "exito", parent)


def mostrar_fallo_proximidad(estado, parent=None, al_crear=None):
    if estado == ESTADO_NO_DETECTADO:
        mostrar_error(
            "Carrito no detectado",
            "No fue posible detectar la señal del carrito móvil.\n"
            "Verifique que el carrito se encuentre encendido.",
            parent=parent,
            al_crear=al_crear,
            texto_destacado="SOLICITE APOYO AL RESPONSABLE"
        )
        return

    if estado == ESTADO_ERROR:
        mostrar_error(
            "Error de verificación",
            "No fue posible verificar la ubicación del equipo.\n"
            "Inténtelo nuevamente si el problema continúa.",
            parent=parent,
            al_crear=al_crear,
            texto_destacado="SOLICITE APOYO AL RESPONSABLE"
        )
        return

    mostrar_error(
        "Ubicación no validada",
        "Para continuar, la laptop debe encontrarse dentro del área del carrito móvil.\n"
        "Acérquese al carrito e inténtelo nuevamente.",
        parent=parent,
        al_crear=al_crear,
        texto_destacado="SI NECESITA AYUDA, SOLICITE APOYO AL RESPONSABLE"
    )


def registrar_uso_contingencia(tipo_operacion, matricula):
    registro_contingencias_sesion.append({
        "fecha_hora": datetime.now(zona_horaria).isoformat(timespec="seconds"),
        "evento": "CONTINGENCIA_ADMIN",
        "operacion": tipo_operacion,
        "matricula": matricula or "NO_DISPONIBLE",
        "equipo": socket.gethostname()
    })
    registro = registro_contingencias_sesion[-1]
    evento = (
        f"{registro['fecha_hora']}\t{registro['evento']}\t{registro['operacion']}\t"
        f"matricula={registro['matricula']}\tequipo={registro['equipo']}\n"
    )

    if MODO_PRUEBA:
        print(evento.rstrip())
        return

    try:
        with open(ARCHIVO_LOG_CONTINGENCIA, "a", encoding="utf-8") as archivo:
            archivo.write(evento)
    except OSError as error:
        print(f"[CONTINGENCIA_ADMIN] No se pudo escribir el registro local: {error}")


def abrir_acceso_administrativo(event=None, parent=None):
    global dialogo_admin_abierto

    if dialogo_admin_abierto:
        return "break"

    dialogo_admin_abierto = True
    parent = parent or ventana

    dialogo = tk.Toplevel(parent)
    dialogo.title("Cierre administrativo")
    dialogo.resizable(False, False)
    dialogo.configure(bg=COLOR_TARJETA)
    dialogo.transient(parent)
    dialogo.attributes("-topmost", True)
    dialogo.withdraw()

    contenido = tk.Frame(dialogo, bg=COLOR_TARJETA, padx=30, pady=24)
    contenido.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        contenido,
        text="Cierre administrativo",
        font=FUENTE_SUBTITULO,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 10))

    mensaje_label = tk.Label(
        contenido,
        text="Introduzca la contraseña administrativa para cerrar el sistema.",
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    )
    mensaje_label.pack(pady=(0, 14))

    tk.Label(
        contenido,
        text="Contraseña:",
        font=FUENTE_PEQ_BOLD,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(anchor="w")

    clave_var = tk.StringVar()
    entrada_clave = tk.Entry(
        contenido,
        textvariable=clave_var,
        show="*",
        justify=tk.CENTER,
        bd=0,
        width=28,
        **estilo_entrada()
    )
    entrada_clave.pack(fill=tk.X, ipady=9, pady=(6, 0))

    estado_intentos = tk.Label(
        contenido,
        text="",
        font=FUENTE_PEQ,
        fg=COLOR_ERROR,
        bg=COLOR_TARJETA
    )
    estado_intentos.pack(pady=(8, 0))

    botones = tk.Frame(contenido, bg=COLOR_TARJETA)
    botones.pack(fill=tk.X, pady=(18, 0))
    intentos_incorrectos = {"cantidad": 0}

    def cerrar():
        global dialogo_admin_abierto
        try:
            dialogo.grab_release()
        except tk.TclError:
            pass
        try:
            dialogo.destroy()
        except tk.TclError:
            pass
        dialogo_admin_abierto = False

    def cerrar_sistema():
        global contexto_contingencia_proximidad
        global confirmacion_contingencia_abierta
        global ventana_entrega_activa

        cerrar()
        contexto_contingencia_proximidad = None
        confirmacion_contingencia_abierta = False
        ventana_entrega_activa = None
        cerrar_aplicacion()

    def autenticar():

        clave = clave_var.get()
        clave_var.set("")
        hash_ingresado = hashlib.sha256(clave.encode("utf-8")).hexdigest()
        clave = ""

        if hmac.compare_digest(hash_ingresado, HASH_CLAVE_ADMIN):
            cerrar_sistema()
            return

        intentos_incorrectos["cantidad"] += 1
        if intentos_incorrectos["cantidad"] >= 3:
            cerrar()
            return

        restantes = 3 - intentos_incorrectos["cantidad"]
        estado_intentos.config(
            text=f"Clave incorrecta. Intentos restantes: {restantes}."
        )
        entrada_clave.focus_set()

    tk.Button(
        botones,
        text="CANCELAR",
        padx=20,
        pady=8,
        command=cerrar,
        **estilo_boton("secundario")
    ).pack(side=tk.RIGHT, padx=(12, 0))

    tk.Button(
        botones,
        text="CERRAR SISTEMA",
        padx=20,
        pady=8,
        command=autenticar,
        **estilo_boton("error")
    ).pack(side=tk.RIGHT)

    dialogo.protocol("WM_DELETE_WINDOW", cerrar)
    dialogo.bind("<Return>", lambda _evento: autenticar())
    dialogo.bind("<Escape>", lambda _evento: cerrar())
    _centrar_dialogo(dialogo, parent, ancho_minimo=460, alto_minimo=285)
    dialogo.grab_set()
    entrada_clave.focus_set()
    return "break"


def iniciar_contexto_contingencia(estado, tipo_operacion, matricula, parent):
    global contexto_contingencia_proximidad
    contexto_contingencia_proximidad = {
        "estado": estado,
        "operacion": tipo_operacion,
        "matricula": matricula,
        "parent": parent,
        "autorizado": False,
        "cerrar_dialogo_fallo": None
    }


def registrar_dialogo_fallo_contingencia(_dialogo, cerrar_dialogo):
    if contexto_contingencia_proximidad is not None:
        contexto_contingencia_proximidad["cerrar_dialogo_fallo"] = cerrar_dialogo


def finalizar_contexto_contingencia():
    global contexto_contingencia_proximidad
    autorizado = bool(
        contexto_contingencia_proximidad
        and contexto_contingencia_proximidad["autorizado"]
    )
    contexto_contingencia_proximidad = None
    return autorizado


def manejar_atajo_administrativo(event=None):
    global confirmacion_contingencia_abierta

    contexto = contexto_contingencia_proximidad
    if contexto and contexto["estado"] in (ESTADO_NO_DETECTADO, ESTADO_ERROR):
        if confirmacion_contingencia_abierta:
            return "break"

        confirmacion_contingencia_abierta = True
        try:
            continuar = mostrar_confirmacion_personalizada(
                "Continuar sin verificación",
                "No fue posible detectar el carrito.\n"
                "¿Desea continuar con el proceso sin validar la proximidad?",
                parent=contexto["parent"]
            )
            if continuar:
                contexto["autorizado"] = True
                registrar_uso_contingencia(
                    contexto["operacion"],
                    contexto["matricula"]
                )
            cerrar_dialogo_fallo = contexto.get("cerrar_dialogo_fallo")
            if cerrar_dialogo_fallo:
                cerrar_dialogo_fallo()
        finally:
            confirmacion_contingencia_abierta = False
        return "break"

    try:
        widget_origen = getattr(event, "widget", None) if event else None
        widget_con_foco = widget_origen or ventana.focus_get()
        if widget_con_foco:
            ventana_con_foco = widget_con_foco.winfo_toplevel()
            if ventana_con_foco == ventana:
                return abrir_acceso_administrativo(event, parent=ventana)
            if (
                ventana_entrega_activa is not None
                and ventana_entrega_activa.winfo_exists()
                and ventana_con_foco == ventana_entrega_activa
            ):
                return abrir_acceso_administrativo(
                    event,
                    parent=ventana_entrega_activa
                )
    except tk.TclError:
        pass

    return "break"


def crear_modal_verificacion_ubicacion(
    parent,
    mensaje="Asegúrate de estar cerca del carrito.",
    titulo="Verificando ubicación del equipo",
    estado="Intentando 1 de 3..."
):
    dialogo = tk.Toplevel(parent)
    dialogo.title(titulo)
    dialogo.resizable(False, False)
    dialogo.configure(bg=COLOR_TARJETA)
    dialogo.transient(parent)
    dialogo.attributes("-topmost", True)
    dialogo.protocol("WM_DELETE_WINDOW", lambda: None)
    dialogo.withdraw()

    contenido = tk.Frame(dialogo, bg=COLOR_TARJETA, padx=30, pady=24)
    contenido.pack(fill=tk.BOTH, expand=True)

    if TEMA_INTERFAZ == "moderno":
        indicador = tk.Canvas(
            contenido, width=52, height=52, bg=COLOR_TARJETA,
            highlightthickness=0
        )
        indicador.pack(pady=(0, 12))
        indicador.create_oval(4, 4, 48, 48, fill=COLORES["info_suave"], outline="")
        indicador.create_text(26, 26, text="⌖", fill=COLOR_PRIMARIO,
                              font=("Segoe UI Symbol", 19, "bold"))

    titulo_label = tk.Label(
        contenido,
        text=titulo,
        font=FUENTE_SUBTITULO,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    )
    titulo_label.pack(pady=(0, 14))

    mensaje_label = tk.Label(
        contenido,
        text=mensaje,
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA,
        justify=tk.CENTER,
        wraplength=350
    )
    mensaje_label.pack()

    estado_intento = tk.Label(
        contenido,
        text=estado,
        font=FUENTE_CUERPO_BOLD,
        fg=COLOR_PRIMARIO,
        bg=(COLORES["info_suave"] if TEMA_INTERFAZ == "moderno" else COLOR_TARJETA)
    )
    estado_intento.pack(
        pady=(18, 0),
        ipadx=(16 if TEMA_INTERFAZ == "moderno" else 0),
        ipady=(7 if TEMA_INTERFAZ == "moderno" else 0)
    )

    _centrar_dialogo(
        dialogo,
        parent,
        ancho_minimo=TAMANOS["dialogo_ancho"],
        alto_minimo=(260 if TEMA_INTERFAZ == "moderno" else 210)
    )
    dialogo.grab_set()
    dialogo.focus_force()
    return dialogo, titulo_label, mensaje_label, estado_intento


def actualizar_modal_progreso(dialogo, titulo_label, mensaje_label, estado_label,
                              titulo, mensaje, estado):
    dialogo.title(titulo)
    titulo_label.config(text=titulo)
    mensaje_label.config(text=mensaje)
    estado_label.config(text=estado)


def convertir_modal_en_confirmacion(
    dialogo,
    titulo_label,
    mensaje_label,
    estado_label,
    nombre,
    al_confirmar,
    al_cancelar
):
    activar_timeout_inactividad()
    registrar_ventana_flujo_alumno(dialogo)
    dialogo.title("Confirmación")
    titulo_label.config(text="Confirmación")
    mensaje_label.config(text=f"¿Eres {nombre}?")
    estado_label.pack_forget()

    botones = tk.Frame(estado_label.master, bg=COLOR_TARJETA)
    botones.pack(fill=tk.X, pady=(22, 0))
    respondido = {"valor": False}

    def responder(confirmado):
        if respondido["valor"]:
            return
        respondido["valor"] = True
        cerrar_modal_verificacion_ubicacion(dialogo)
        if confirmado:
            al_confirmar()
        else:
            al_cancelar()

    tk.Button(
        botones,
        text="NO",
        padx=26,
        pady=8,
        command=lambda: responder(False),
        **estilo_boton("secundario")
    ).pack(side=tk.RIGHT, padx=(12, 0))

    tk.Button(
        botones,
        text="SÍ",
        padx=26,
        pady=8,
        command=lambda: responder(True),
        **estilo_boton("principal")
    ).pack(side=tk.RIGHT)

    dialogo.protocol("WM_DELETE_WINDOW", lambda: responder(False))
    dialogo.bind("<Return>", lambda _evento: responder(True))
    dialogo.bind("<Escape>", lambda _evento: responder(False))
    _centrar_dialogo(dialogo, dialogo.master, ancho_minimo=430, alto_minimo=210)
    dialogo.focus_force()


def cerrar_modal_verificacion_ubicacion(dialogo):
    try:
        dialogo.grab_release()
    except tk.TclError:
        pass
    try:
        dialogo.destroy()
    except tk.TclError:
        pass


def centrar_ventana(ventana, ancho, alto):
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


# --- FUNCIONES BASE (SIN CAMBIOS) ---
def apagar_windows():
    if threading.current_thread() is not threading.main_thread():
        ventana.after(0, apagar_windows)
        return

    if MODO_PRUEBA:
        mostrar_exito(
            "Modo de prueba",
            "La devolución terminó correctamente. En funcionamiento normal, "
            "la laptop se apagaría en este momento."
        )
        cerrar_aplicacion()
        return

    os.system("shutdown /s /t 3")


def contar_usos_alumno(matricula):
    """
    Cuenta cuántas veces el alumno ha usado el sistema
    con base en la hoja 'Registros'.
    """
    try:
        if hoja_registros is None:
            return 0

        matriculas = hoja_registros.col_values(1)
        return matriculas.count(matricula)

    except Exception as e:
        print(f"Error al contar usos del alumno: {e}")
        return 0


def _calcular_alto_aviso(alto_requerido, alto_pantalla):
    return min(max(alto_requerido + 16, 230), alto_pantalla - 80)


def mostrar_ventana_control_unificada(
    matricula,
    nombre,
    tiene_pendiente,
    fecha_entrada,
    laptop_id,
    no_entregas,
    estado
):
    ventana_ctrl = tk.Toplevel(ventana)
    registrar_ventana_flujo_alumno(ventana_ctrl)
    ventana_ctrl.title("Aviso importante")
    ventana_ctrl.resizable(False, False)
    ventana_ctrl.configure(bg="#ffffff")
    
    # Aplicar estilo moderno
    ventana_ctrl.configure(bg=COLOR_FONDO)
    
    # Hacerla modal respecto a la ventana principal
    ventana_ctrl.transient(ventana)
    ventana_ctrl.grab_set()

    ancho = 500
    frame = tk.Frame(ventana_ctrl, bg=COLOR_TARJETA)
    frame.pack(fill=tk.BOTH, expand=True)

    contenido = tk.Frame(frame, bg=COLOR_TARJETA, padx=30, pady=12)
    contenido.pack(fill=tk.X)

    boton_frame = tk.Frame(frame, bg=COLOR_TARJETA, padx=40, pady=5)
    boton_frame.pack(side=tk.BOTTOM, fill=tk.X)


    # USUARIO SANCIONADO
    if estado == "SANCIONADO":
        titulo = "USUARIO SANCIONADO"
        color = COLOR_ERROR
        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n\n"
            "Has excedido el número permitido.\n\n"
            "Acude con el administrador."
        )
    else:
        # AVISO GENERAL
        titulo = "AVISO IMPORTANTE"
        color = COLOR_ADVERTENCIA

        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n"
            f"Estado: {estado}\n\n"
            "Recuerda entregar correctamente la laptop al finalizar su uso.\n\n"
            "Al alcanzar el límite permitido, el acceso será bloqueado."
        )


        if tiene_pendiente:
            mensaje = (
                "Entrega pendiente detectada\n\n"
                f"Laptop: {laptop_id}\n"
                f"Fecha: {formatear_fecha(fecha_entrada)}\n\n"
            ) + mensaje

    # Icono

    tk.Label(
        contenido,
        text=titulo,
        font=("Segoe UI", 16, "bold"),
        fg=color,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 10))

    tk.Label(
        contenido,
        text=mensaje,
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA,
        justify=tk.CENTER,
        wraplength=420
    ).pack(pady=(8, 6), fill=tk.X)


    
    tk.Button(
        boton_frame,
        text="ENTENDIDO",
        font=("Segoe UI", 13, "bold"),   # ? más equilibrado
        bg=COLOR_PRIMARIO,
        fg="white",
        bd=0,
        cursor="hand2",
        activebackground=COLOR_SECUNDARIO,
        activeforeground="white",
        relief="flat",
        command=ventana_ctrl.destroy
    ).pack(
        fill=tk.X,
        padx=90,        # ? más angosto
        pady=(8, 5),
        ipady=6         # ? altura fina y controlada
    )





    ventana_ctrl.update_idletasks()
    pantalla_ancho = ventana_ctrl.winfo_screenwidth()
    pantalla_alto = ventana_ctrl.winfo_screenheight()
    alto = _calcular_alto_aviso(ventana_ctrl.winfo_reqheight(), pantalla_alto)
    x = (pantalla_ancho - ancho) // 2
    y = (pantalla_alto - alto) // 2
    ventana_ctrl.geometry(f"{ancho}x{alto}+{x}+{y}")

    # ESPERAR a que el usuario cierre la ventana
    ventana_ctrl.wait_window()

def matricula_exenta_no_entrega(matricula):
    return str(matricula).strip() in MATRICULAS_EXENTAS_NO_ENTREGA


def incrementar_no_entregas(matricula, contador_objetivo):
    """
    Lleva No_Entregas hasta un objetivo persistido en la fila de Registros.
    El estado se calcula ÚNICAMENTE por fórmula en Google Sheets.
    """
    try:
        matricula = str(matricula).strip()
        if matricula_exenta_no_entrega(matricula):
            return 0, "EXENTO"

        if hoja_alumnos is None:
            raise ErrorConsultaGoogle("La hoja Alumnos no está disponible")

        matriculas = hoja_alumnos.col_values(1)
        if matricula not in matriculas:
            raise ErrorConsultaGoogle("La matrícula no existe en la hoja Alumnos")

        fila = matriculas.index(matricula) + 1

        valor_actual = hoja_alumnos.cell(fila, 5).value
        no_entregas_actual = int(valor_actual) if valor_actual else 0
        no_entregas_nuevo = max(no_entregas_actual, int(contador_objetivo))

        try:
            hoja_alumnos.update_cell(fila, 5, no_entregas_nuevo)
        except Exception as error_actualizacion:
            print(f"Respuesta ambigua al sincronizar no entregas: {error_actualizacion}")

        valor_confirmado = hoja_alumnos.cell(fila, 5).value
        try:
            contador_confirmado = int(valor_confirmado) if valor_confirmado else 0
        except (TypeError, ValueError) as error:
            raise ErrorConsultaGoogle("No se pudo confirmar No_Entregas") from error
        if contador_confirmado != no_entregas_nuevo:
            raise ErrorConsultaGoogle("No se pudo confirmar No_Entregas")

        # NO TOCAR columna del estado (tiene fórmula)
        return no_entregas_nuevo, "CALCULADO_POR_FORMULA"

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al incrementar no entregas: {e}")
        raise ErrorConsultaGoogle("No se pudo actualizar el control del alumno") from e
    
def _contador_no_entregas_actual(matricula):
    matriculas = hoja_alumnos.col_values(1)
    if matricula not in matriculas:
        raise ErrorConsultaGoogle("La matrícula no existe en la hoja Alumnos")
    fila = matriculas.index(matricula) + 1
    valor = hoja_alumnos.cell(fila, 5).value
    return int(valor) if valor else 0


def _objetivo_estado_intermedio(observacion, prefijo):
    if not observacion.startswith(prefijo + ":"):
        return None
    try:
        return int(observacion.split(":", 1)[1])
    except (TypeError, ValueError):
        raise ErrorConsultaGoogle("Estado intermedio de no entrega inválido")


def _finalizar_observacion(numero_fila, observacion_final):
    try:
        hoja_registros.update_cell(numero_fila, COL_OBSERVACION + 1, observacion_final)
    except Exception as error_actualizacion:
        print(f"Respuesta ambigua al finalizar no entrega: {error_actualizacion}")
    confirmado = hoja_registros.cell(numero_fila, COL_OBSERVACION + 1).value
    if str(confirmado or "").strip() != observacion_final:
        raise ErrorConsultaGoogle("No se pudo confirmar la observación final")


def cerrar_sesion_anterior_y_contar_no_entrega(matricula):
    with bloqueo_local_registro_prestamo():
        return _cerrar_sesion_anterior_y_contar_no_entrega(matricula)


def _cerrar_sesion_anterior_y_contar_no_entrega(matricula):
    """
    Cierra la sesión activa anterior del alumno,
    registra NO ENTREGA y deja evidencia.
    """
    try:
        if hoja_registros is None or hoja_alumnos is None:
            raise ErrorConsultaGoogle("Las hojas de Google no están disponibles")

        registros = hoja_registros.get_all_values()
        hora_actual, _ = obtener_hora_internet()

        for i in reversed(range(len(registros))):
            fila = registros[i]

            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                observacion = fila[COL_OBSERVACION].strip()
                if observacion == "CIERRE_AUTOMATICO_POR_NUEVA_SESION":
                    return True

                objetivo_pendiente = _objetivo_estado_intermedio(
                    observacion, "PROCESANDO_CIERRE_AUTOMATICO"
                )
                if objetivo_pendiente is not None:
                    if not matricula_exenta_no_entrega(matricula):
                        incrementar_no_entregas(matricula, objetivo_pendiente)
                    _finalizar_observacion(
                        i + 1, "CIERRE_AUTOMATICO_POR_NUEVA_SESION"
                    )
                    return True

                if (
                    fila[COL_HORA_SALIDA].strip() == ""
                    and observacion_es_normal(fila[COL_OBSERVACION])
                ):
                    numero_fila = i + 1
                    objetivo = _contador_no_entregas_actual(matricula) + 1
                    estado_actual = hoja_registros.get(
                        f"E{numero_fila}:G{numero_fila}"
                    )
                    if not estado_actual or len(estado_actual[0]) <= 2:
                        raise ErrorConsultaGoogle("No se pudo releer la sesión anterior")
                    salida_actual = estado_actual[0][0].strip()
                    observacion_actual = estado_actual[0][2].strip()
                    if observacion_actual == "CIERRE_AUTOMATICO_POR_NUEVA_SESION":
                        return True
                    objetivo_existente = _objetivo_estado_intermedio(
                        observacion_actual, "PROCESANDO_CIERRE_AUTOMATICO"
                    )
                    if objetivo_existente is not None:
                        if not matricula_exenta_no_entrega(matricula):
                            incrementar_no_entregas(matricula, objetivo_existente)
                        _finalizar_observacion(
                            numero_fila, "CIERRE_AUTOMATICO_POR_NUEVA_SESION"
                        )
                        return True
                    if salida_actual or not observacion_es_normal(observacion_actual):
                        return False

                    estado_intermedio = f"PROCESANDO_CIERRE_AUTOMATICO:{objetivo}"
                    try:
                        hoja_registros.batch_update([
                            {"range": f"E{numero_fila}", "values": [[hora_actual]]},
                            {"range": f"G{numero_fila}", "values": [[estado_intermedio]]}
                        ])
                    except Exception as error_marcado:
                        print(f"Respuesta ambigua al cerrar sesión anterior: {error_marcado}")
                    verificacion = hoja_registros.get(f"E{numero_fila}:G{numero_fila}")
                    if (
                        not verificacion
                        or len(verificacion[0]) <= 2
                        or verificacion[0][0].strip() != hora_actual
                        or verificacion[0][2].strip() not in (
                            estado_intermedio,
                            "CIERRE_AUTOMATICO_POR_NUEVA_SESION",
                        )
                    ):
                        raise ErrorConsultaGoogle("No se pudo verificar el cierre automático")

                    if verificacion[0][2].strip() == "CIERRE_AUTOMATICO_POR_NUEVA_SESION":
                        return True
                    if not matricula_exenta_no_entrega(matricula):
                        incrementar_no_entregas(matricula, objetivo)
                    _finalizar_observacion(
                        numero_fila, "CIERRE_AUTOMATICO_POR_NUEVA_SESION"
                    )
                    return True
                break

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al cerrar sesión anterior: {e}")
        raise ErrorConsultaGoogle("No se pudo cerrar la sesión anterior") from e

    return False


def obtener_control_alumno(matricula):
    """
    Obtiene el número de no entregas y el estado del alumno
    desde la hoja 'Alumnos'.
    """
    try:
        if hoja_alumnos is None:
            raise ErrorConsultaGoogle("La hoja Alumnos no está disponible")

        matriculas = hoja_alumnos.col_values(1)
        if matricula in matriculas:
            fila = matriculas.index(matricula) + 1

            no_entregas = hoja_alumnos.cell(fila, 5).value
            estado = hoja_alumnos.cell(fila, 7).value

            no_entregas = int(no_entregas) if no_entregas else 0
            estado = normalizar_estado_alumno(estado, lectura_exitosa=True)

            return no_entregas, estado

        raise ErrorConsultaGoogle("La matrícula no existe en la hoja Alumnos")

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al obtener control del alumno: {e}")
        raise ErrorConsultaGoogle("No se pudo consultar el control del alumno") from e

def mostrar_instrucciones_iniciales(matricula=""):
    ventana_info = tk.Toplevel(ventana)
    ventana_info.title("Instrucciones de Uso")
    ventana_info.resizable(False, False)
    ventana_info.configure(bg=COLOR_FONDO)

    # Modal correcta (SIN topmost)
    ventana_info.transient(ventana)
    ventana_info.grab_set()
    ventana_info.focus_force()

    ancho = 560
    alto = 400

    x = (ventana_info.winfo_screenwidth() - ancho) // 2
    y = (ventana_info.winfo_screenheight() - alto) // 2
    ventana_info.geometry(f"{ancho}x{alto}+{x}+{y}")

    frame = tk.Frame(
        ventana_info,
        bg=COLOR_TARJETA,
        padx=35,
        pady=30
    )
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        frame,
        text="INSTRUCCIONES IMPORTANTES",
        font=("Segoe UI", 16, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 7))

    texto = (
    "• La laptop es responsabilidad del alumno durante su uso.\n"
    "• Cualquier daño o falla debe reportarse de inmediato.\n"
    "• NO cierres el sistema ni apagues la computadora manualmente.\n"
    "• Siempre usa el botón “ENTREGAR Y APAGAR”.\n"
    "• No entregar la laptop genera NO ENTREGA y posibles sanciones.\n"
    )

    texto_widget = tk.Text(
    frame,
    font=FUENTE_CUERPO,
    fg=COLOR_TEXTO_SECUNDARIO,
    bg=COLOR_TARJETA,
    wrap=tk.WORD,
    relief=tk.FLAT,
    height=8,        # clave
    padx=20,
    pady=4
)
    texto_widget.pack(pady=6, fill=tk.X)

    texto_widget.insert(tk.END, texto)
    texto_widget.config(state=tk.DISABLED)




    tk.Button(
        frame,
        text="ENTENDIDO",
        font=FUENTE_BOTON,
        bg=COLOR_EXITO,
        fg="white",
        activebackground="#2ecc71",
        padx=35,
        pady=12,
        bd=0,
        cursor="hand2",
        command=ventana_info.destroy
    ).pack(pady=20)

def bloquear_alt_f4(event):
    """
    Bloquea la combinación Alt + F4 para usuarios normales.
    """
    return "break"

def cargar_logo(ruta_imagen, ancho, alto):
    """Carga y redimensiona el logo"""
    if not PIL_DISPONIBLE:
        return None
    try:
        imagen = Image.open(ruta_imagen)
        imagen = imagen.resize((ancho, alto), Image.LANCZOS)
        return ImageTk.PhotoImage(imagen)
    except Exception as e:
        print(f"Error cargando logo: {e}")
        return None


def crear_fondo_marca_agua_login(ruta_imagen, ancho, alto):
    """Genera una marca de agua ligera para el fondo exterior del login."""
    if not PIL_DISPONIBLE or not FONDO_MARCA_AGUA_LOGIN:
        return None

    try:
        fondo = Image.new("RGB", (ancho, alto), COLOR_FONDO)
        logo = Image.open(ruta_imagen).convert("RGBA")

        mascara_recorte = ImageChops.multiply(
            ImageOps.invert(ImageOps.grayscale(logo)),
            logo.getchannel("A")
        )
        limites_logo = mascara_recorte.getbbox()
        if limites_logo:
            logo = logo.crop(limites_logo)

        logo.thumbnail((72, 58), Image.LANCZOS)

        escala_grises = ImageOps.grayscale(logo)
        mascara_visible = ImageChops.multiply(
            ImageOps.invert(escala_grises),
            logo.getchannel("A")
        )
        mascara = mascara_visible.point(
            lambda valor: int(valor * 0.17)
        )
        marca = Image.new("RGBA", logo.size, (151, 180, 201, 0))
        marca.putalpha(mascara)
        marca = marca.rotate(-18, expand=True, resample=Image.BICUBIC)

        paso_x, paso_y = 145, 118
        for fila, y in enumerate(range(-20, alto + paso_y, paso_y)):
            desplazamiento = 70 if fila % 2 else 0
            for x in range(-70 + desplazamiento, ancho + paso_x, paso_x):
                fondo.paste(marca, (x, y), marca)

        return ImageTk.PhotoImage(fondo)
    except Exception as e:
        print(f"Error creando marca de agua del login: {e}")
        return None

def verificar_internet():
    urls_prueba = [
        "https://www.google.com",
        "https://sheets.googleapis.com",
        "https://www.microsoft.com"
    ]

    for url in urls_prueba:
        try:
            respuesta = requests.get(
                url,
                timeout=2,
                verify=certifi.where(),
                headers={"User-Agent": "Mozilla/5.0"},
                allow_redirects=True
            )
            if respuesta.status_code < 500:
                return True
                
        except Exception as e:
            pass

    return False


def registrar_log_actualizacion(mensaje):
    marca_tiempo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ruta_log = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        ARCHIVO_LOG_ACTUALIZACION
    )
    try:
        with open(ruta_log, "a", encoding="utf-8") as archivo:
            archivo.write(f"[{marca_tiempo}] [ACTUALIZACION] {mensaje}\n")
    except OSError:
        pass


def ejecutar_git_segundo_plano(argumentos, timeout=15):
    directorio = os.path.dirname(os.path.abspath(__file__))
    entorno = os.environ.copy()
    entorno["GIT_TERMINAL_PROMPT"] = "0"
    entorno["GCM_INTERACTIVE"] = "Never"
    return subprocess.run(
        ["git", *argumentos],
        cwd=directorio,
        stdin=subprocess.DEVNULL,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=entorno,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        check=False
    )


def comprobar_actualizacion_segundo_plano():
    if not actualizacion_lock.acquire(blocking=False):
        return

    try:
        registrar_log_actualizacion("comprobando...")
        try:
            respuesta = requests.get(
                "https://github.com",
                timeout=3,
                verify=certifi.where(),
                headers={"User-Agent": "RegistroLaptop-Updater"}
            )
            if respuesta.status_code >= 500:
                registrar_log_actualizacion("sin Internet")
                return
        except requests.RequestException:
            registrar_log_actualizacion("sin Internet")
            return

        resultado_fetch = ejecutar_git_segundo_plano(
            ["fetch", "origin", "main", "--quiet"],
            timeout=15
        )
        if resultado_fetch.returncode != 0:
            detalle = resultado_fetch.stderr or resultado_fetch.stdout
            detalle = " ".join((detalle or "error desconocido").split())[:300]
            registrar_log_actualizacion(f"error en fetch: {detalle}")
            return

        pendientes = ejecutar_git_segundo_plano(
            ["rev-list", "--count", "HEAD..origin/main"],
            timeout=5
        )
        if pendientes.returncode != 0:
            registrar_log_actualizacion("error comprobando la versión remota")
            return

        try:
            cantidad_pendiente = int(pendientes.stdout.strip() or "0")
        except ValueError:
            registrar_log_actualizacion("error interpretando la versión remota")
            return

        if cantidad_pendiente == 0:
            registrar_log_actualizacion("sin cambios")
            return

        directorio = os.path.dirname(os.path.abspath(__file__))
        ruta_marker = os.path.join(directorio, ".actualizacion_pendiente")
        ruta_temporal = ruta_marker + ".tmp"
        with open(ruta_temporal, "w", encoding="ascii") as archivo:
            archivo.write("pendiente\n")
        os.replace(ruta_temporal, ruta_marker)
        registrar_log_actualizacion(
            "nueva versión detectada; actualización marcada como pendiente"
        )
    except FileNotFoundError:
        registrar_log_actualizacion("error: Git no está disponible")
    except subprocess.TimeoutExpired:
        registrar_log_actualizacion("error: tiempo agotado al comprobar")
    except Exception as error:
        registrar_log_actualizacion(f"error inesperado: {error}")
    finally:
        actualizacion_lock.release()


def iniciar_actualizacion_en_segundo_plano():
    threading.Thread(
        target=comprobar_actualizacion_segundo_plano,
        daemon=True
    ).start()


def conectar_google_sheets(internet_verificado=False):
    global hoja_alumnos, hoja_registros

    if not internet_verificado and not verificar_internet():
        cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        return False

    with conexion_google_lock:
        if hoja_alumnos is not None and hoja_registros is not None:
            if verificar_conexion_base_datos():
                return True
            hoja_alumnos = None
            hoja_registros = None

        cambiar_estado("Verificando conexión...", COLOR_ADVERTENCIA)
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                RUTA_CREDENCIALES, scope
            )

            client = gspread.authorize(creds)
            client.http_client.set_timeout((3, 10))

            sheet = client.open("Control de Laptops")

            hoja_alumnos = sheet.worksheet("Alumnos")
            hoja_registros = sheet.worksheet("Registros")

            cambiar_estado("Conectado", COLOR_EXITO)
            return True
        except FileNotFoundError:
            hoja_alumnos = None
            hoja_registros = None
            cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
            return False
        except Exception as e:
            hoja_alumnos = None
            hoja_registros = None
            cambiar_estado("Error en la conexión", COLOR_ERROR)
            return False


def verificar_conexion_base_datos():
    """Comprueba mediante una lectura remota que Google Sheets responde."""
    if hoja_alumnos is None or hoja_registros is None:
        return False
    try:
        hoja_alumnos.acell("A1")
        return True
    except Exception as error:
        print(f"Error al verificar Google Sheets: {error}")
        return False


def verificar_sheets_con_reintentos(max_intentos=3, callback_intento=None):
    """Verifica Sheets remotamente con pausas breves, sin realizar escrituras."""
    pausas = (0, 1, 2)
    for intento in range(1, max_intentos + 1):
        if intento > 1:
            time.sleep(pausas[min(intento - 1, len(pausas) - 1)])
        if callback_intento:
            callback_intento(intento, max_intentos)

        conectado = verificar_conexion_base_datos()
        if not conectado:
            conectado = conectar_google_sheets(internet_verificado=True)
            if conectado:
                conectado = verificar_conexion_base_datos()

        registrar_evento_tecnico(
            f"SHEETS_CONEXION intento={intento} resultado={'OK' if conectado else 'ERROR'}"
        )
        if conectado:
            return True
    return False

def verificar_conexion_periodicamente():
    global verificar_conexion_activo, aviso_internet
    
    while verificar_conexion_activo:
        if verificar_internet():
            if hoja_alumnos is None or hoja_registros is None:
                if conectar_google_sheets(internet_verificado=True):
                    # Cerrar aviso de internet si está abierto
                    ejecutar_en_ui(cerrar_aviso_internet)
                    
            elif verificar_conexion_base_datos():
                cambiar_estado("Conectado", COLOR_EXITO)
            else:
                cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        else:
            cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        
        time.sleep(2)

def detener_verificacion_conexion():
    global verificar_conexion_activo
    verificar_conexion_activo = False


def cerrar_aplicacion():
    """Finaliza Tkinter y permite que termine el proceso principal."""
    global cierre_aplicacion_en_curso
    if cierre_aplicacion_en_curso:
        return

    cierre_aplicacion_en_curso = True
    registrar_evento_tecnico("CIERRE_APLICACION_INICIADO")
    detener_verificacion_conexion()
    cancelar_timeout_inactividad()
    try:
        if ventana.winfo_exists():
            ventana.destroy()
    except (tk.TclError, NameError):
        pass
    registrar_evento_tecnico("PROCESO_FINALIZANDO")

def ejecutar_en_ui(func, *args):
    try:
        if threading.current_thread() is threading.main_thread():
            func(*args)
        else:
            ventana.after(0, lambda: func(*args))
    except:
        pass


def registrar_ventana_flujo_alumno(dialogo):
    if not timeout_inactividad_activo:
        return
    ventanas_flujo_alumno.add(dialogo)

    def olvidar_dialogo(event):
        if event.widget == dialogo:
            ventanas_flujo_alumno.discard(dialogo)

    dialogo.bind("<Destroy>", olvidar_dialogo, add="+")


def cancelar_timeout_inactividad():
    global timeout_inactividad_id, timeout_inactividad_activo
    timeout_inactividad_activo = False
    if timeout_inactividad_id is not None:
        try:
            ventana.after_cancel(timeout_inactividad_id)
        except (tk.TclError, ValueError):
            pass
        timeout_inactividad_id = None


def reiniciar_timeout_inactividad(event=None):
    global timeout_inactividad_id
    if not timeout_inactividad_activo:
        return
    if timeout_inactividad_id is not None:
        try:
            ventana.after_cancel(timeout_inactividad_id)
        except (tk.TclError, ValueError):
            pass
    timeout_inactividad_id = ventana.after(
        TIEMPO_INACTIVIDAD_MS,
        procesar_timeout_inactividad
    )


def activar_timeout_inactividad():
    global timeout_inactividad_activo, flujo_cancelado_por_inactividad
    timeout_inactividad_activo = True
    flujo_cancelado_por_inactividad = False
    reiniciar_timeout_inactividad()


def procesar_timeout_inactividad():
    global timeout_inactividad_id, timeout_inactividad_activo
    global flujo_cancelado_por_inactividad, procesando_sesion
    global contexto_contingencia_proximidad, confirmacion_contingencia_abierta

    if not timeout_inactividad_activo:
        return

    timeout_inactividad_id = None
    timeout_inactividad_activo = False
    flujo_cancelado_por_inactividad = True
    contexto_contingencia_proximidad = None
    confirmacion_contingencia_abierta = False
    procesando_sesion = False

    for dialogo in list(ventanas_flujo_alumno):
        try:
            if dialogo.winfo_exists():
                dialogo.destroy()
        except tk.TclError:
            pass
    ventanas_flujo_alumno.clear()

    try:
        entrada.delete(0, tk.END)
        acepta_estado_equipo.set(False)
        btn_entrar.config(state="normal", text="INGRESAR AL SISTEMA")
        cambiar_estado(
            "Proceso cancelado por inactividad",
            COLOR_ADVERTENCIA
        )
        ventana.deiconify()
        ventana.lift()
        ventana.after(100, entrada.focus_set)
    except (tk.TclError, NameError):
        pass


def registrar_actividad_usuario(event=None):
    reiniciar_timeout_inactividad()


def cerrar_aviso_internet():
    global aviso_internet
    try:
        if aviso_internet and aviso_internet.winfo_exists():
            aviso_internet.destroy()
        aviso_internet = None
    except:
        aviso_internet = None
def obtener_hora_internet():
    try:
        response = requests.get(
            "http://worldtimeapi.org/api/timezone/America/Chihuahua",
            timeout=5,
            verify=certifi.where(),
            headers={"User-Agent": "Mozilla/5.0"}
        )
        hora_utc = datetime.fromisoformat(response.json()["datetime"].split(".")[0])
        return zona_horaria.localize(hora_utc).strftime("%H:%M:%S"), zona_horaria.localize(hora_utc).strftime("%Y-%m-%d")
    except:
        ahora = datetime.now(zona_horaria)
        return ahora.strftime("%H:%M:%S"), ahora.strftime("%Y-%m-%d")

def buscar_nombre(matricula):
    try:
        if hoja_alumnos is None:
            return None
        celdas = hoja_alumnos.col_values(1)
        if matricula in celdas:
            row = celdas.index(matricula) + 1
            return hoja_alumnos.cell(row, 2).value
    except Exception as e:
        print(f"Error al buscar nombre: {e}")
        return None
    return None

def buscar_rol(matricula):
    """
    Obtiene el rol del usuario desde Google Sheets
    Columna C
    """
    try:
        if hoja_alumnos is None:
            return "ALUMNO"

        celdas = hoja_alumnos.col_values(1)
        if matricula in celdas:
            row = celdas.index(matricula) + 1
            rol = hoja_alumnos.cell(row, 3).value
            return rol.strip().upper() if rol else "ALUMNO"
    except Exception as e:
        print(f"Error al buscar rol: {e}")
        return "ALUMNO"


def buscar_curp(matricula):
    """
    Obtiene la CURP completa desde Google Sheets
    Columna D
    """
    try:
        if hoja_alumnos is None:
            return None

        celdas = hoja_alumnos.col_values(1)
        if matricula in celdas:
            row = celdas.index(matricula) + 1
            return hoja_alumnos.cell(row, 4).value
    except Exception as e:
        print(f"Error al buscar CURP: {e}")
        return None


def validar_curp_ultimos_2(curp_real, curp_ingresada):
    if not curp_real or len(curp_real) < 2:
        return False
    return curp_real[-2:] == curp_ingresada.strip()


def pedir_curp_ultimos_2(parent):
    activar_timeout_inactividad()
    ventana_curp = tk.Toplevel(parent)
    registrar_ventana_flujo_alumno(ventana_curp)
    ventana_curp.title("Validación de identidad")
    ventana_curp.resizable(False, False)
    ventana_curp.configure(bg=COLOR_FONDO)
    ventana_curp.transient(parent)
    ventana_curp.attributes("-topmost", True)
    ventana_curp.withdraw()

    tarjeta = tk.Frame(
        ventana_curp,
        bg=COLOR_TARJETA,
        highlightthickness=1,
        highlightbackground=COLOR_BORDE,
        padx=34,
        pady=28
    )
    tarjeta.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    tk.Label(
        tarjeta,
        text="Validación de identidad",
        font=("Segoe UI", 17, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 8))

    tk.Label(
        tarjeta,
        text="Por seguridad, ingresa los últimos 2 caracteres de tu CURP.",
        font=FUENTE_CUERPO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA,
        justify=tk.CENTER,
        wraplength=360
    ).pack(pady=(0, 22))

    tk.Label(
        tarjeta,
        text="ÚLTIMOS 2 CARACTERES",
        font=FUENTE_PEQ_BOLD,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(anchor="w")

    resultado = {"valor": None}

    # VARIABLE CONTROLADA
    curp_var = tk.StringVar()

    def controlar_texto(*args):
        valor = curp_var.get().upper()

        # Limitar a 2 caracteres
        if len(valor) > 2:
            valor = valor[:2]

        # Evitar loop infinito
        if curp_var.get() != valor:
            curp_var.set(valor)

    # Vincular control
    curp_var.trace_add("write", controlar_texto)

    estilo_campo_curp = estilo_entrada()
    estilo_campo_curp["font"] = ("Segoe UI", 18, "bold")
    entrada = tk.Entry(
        tarjeta,
        textvariable=curp_var,
        justify="center",
        **estilo_campo_curp
    )
    entrada.pack(fill=tk.X, pady=(7, 20), ipady=10)

    def confirmar(event=None):
        valor = curp_var.get().strip()

        if len(valor) != 2:
            mostrar_error(
                "Dato inválido",
                "Debes ingresar exactamente 2 caracteres.",
                parent=ventana_curp
            )
            entrada.focus_force()
            entrada.selection_range(0, tk.END)
            return

        resultado["valor"] = valor
        ventana_curp.attributes("-topmost", False)
        ventana_curp.destroy()

    entrada.bind("<Return>", confirmar)

    tk.Button(
        tarjeta,
        text="VALIDAR",
        padx=25,
        pady=10,
        command=confirmar,
        **estilo_boton("principal")
    ).pack(fill=tk.X)

    _centrar_dialogo(ventana_curp, parent, ancho_minimo=440, alto_minimo=310)
    ventana_curp.grab_set()
    ventana_curp.lift()

    def enfocar_entrada():
        entrada.focus_force()
        entrada.selection_range(0, tk.END)
        entrada.icursor(tk.END)

    ventana_curp.after_idle(enfocar_entrada)
    ventana_curp.wait_window()
    return resultado["valor"]

def verificar_entrega_pendiente(matricula):
    try:
        if hoja_registros is None:
            raise ErrorConsultaGoogle("La hoja Registros no está disponible")

        registros = hoja_registros.get_all_values()

        for i in reversed(range(len(registros))):
            fila = registros[i]
            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                observacion = fila[COL_OBSERVACION].strip()
                if hora_salida == "" and observacion_es_normal(observacion):
                    fecha_entrada = fila[COL_FECHA]
                    laptop_id = fila[COL_LAPTOP_ID]
                    return True, fecha_entrada, laptop_id

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al verificar entrega pendiente: {e}")
        raise ErrorConsultaGoogle("No se pudo verificar la entrega pendiente") from e

    return False, None, None


def procesar_no_entrega_si_corresponde(matricula):
    with bloqueo_local_registro_prestamo():
        return _procesar_no_entrega_si_corresponde(matricula)


def _procesar_no_entrega_si_corresponde(matricula):
    try:
        if hoja_registros is None or hoja_alumnos is None:
            raise ErrorConsultaGoogle("Las hojas de Google no están disponibles")

        registros = hoja_registros.get_all_values()

        for i in reversed(range(len(registros))):
            fila = registros[i]

            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                observacion = fila[COL_OBSERVACION].strip()

                if observacion == "NO_ENTREGA_CONTADA":
                    break

                objetivo_pendiente = _objetivo_estado_intermedio(
                    observacion, "PROCESANDO_NO_ENTREGA"
                )
                if objetivo_pendiente is not None:
                    if not matricula_exenta_no_entrega(matricula):
                        incrementar_no_entregas(matricula, objetivo_pendiente)
                    _finalizar_observacion(i + 1, "NO_ENTREGA_CONTADA")
                    break

                # NO ENTREGA: existe una sesión pendiente todavía no contabilizada.
                if (
                    hora_salida == ""
                    and observacion_es_normal(observacion)
                    and not matricula_exenta_no_entrega(matricula)
                ):
                    numero_fila = i + 1
                    objetivo = _contador_no_entregas_actual(matricula) + 1
                    estado_actual = hoja_registros.get(
                        f"E{numero_fila}:G{numero_fila}"
                    )
                    if not estado_actual or len(estado_actual[0]) <= 2:
                        raise ErrorConsultaGoogle("No se pudo releer la no entrega")
                    salida_actual = estado_actual[0][0].strip()
                    observacion_actual = estado_actual[0][2].strip()
                    if observacion_actual == "NO_ENTREGA_CONTADA":
                        break
                    objetivo_existente = _objetivo_estado_intermedio(
                        observacion_actual, "PROCESANDO_NO_ENTREGA"
                    )
                    if objetivo_existente is not None:
                        incrementar_no_entregas(matricula, objetivo_existente)
                        _finalizar_observacion(numero_fila, "NO_ENTREGA_CONTADA")
                        break
                    if salida_actual or not observacion_es_normal(observacion_actual):
                        break

                    estado_intermedio = f"PROCESANDO_NO_ENTREGA:{objetivo}"
                    try:
                        hoja_registros.batch_update([
                            {"range": f"E{numero_fila}", "values": [["-----"]]},
                            {"range": f"G{numero_fila}", "values": [[estado_intermedio]]}
                        ])
                    except Exception as error_marcado:
                        print(f"Respuesta ambigua al marcar no entrega: {error_marcado}")
                    verificacion = hoja_registros.get(f"E{numero_fila}:G{numero_fila}")
                    if (
                        not verificacion
                        or len(verificacion[0]) <= 2
                        or verificacion[0][0].strip() != "-----"
                        or verificacion[0][2].strip() not in (
                            estado_intermedio,
                            "NO_ENTREGA_CONTADA",
                        )
                    ):
                        raise ErrorConsultaGoogle("No se pudo verificar el procesamiento de no entrega")
                    if verificacion[0][2].strip() == "NO_ENTREGA_CONTADA":
                        break
                    incrementar_no_entregas(matricula, objetivo)
                    _finalizar_observacion(numero_fila, "NO_ENTREGA_CONTADA")
                break

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al procesar no entrega: {e}")
        raise ErrorConsultaGoogle("No se pudo procesar la no entrega") from e


def sesion_activa_en_esta_laptop(matricula):
    try:
        if hoja_registros is None:
            return False

        laptop_actual = socket.gethostname()
        registros = hoja_registros.get_all_values()

        for fila in reversed(registros):
            if len(fila) > COL_OBSERVACION and fila[COL_MATRICULA] == matricula:
                return (
                    fila[COL_LAPTOP_ID] == laptop_actual and
                    fila[COL_HORA_SALIDA].strip() == "" and
                    observacion_es_normal(fila[COL_OBSERVACION])
                )

    except Exception as e:
        print(f"Error al validar sesión activa local: {e}")

    return False




def _existe_sesion_activa_para_matricula(registros, matricula):
    return any(
        len(fila) > COL_OBSERVACION
        and str(fila[COL_MATRICULA]).strip() == str(matricula).strip()
        and str(fila[COL_HORA_SALIDA]).strip() == ""
        and observacion_es_normal(fila[COL_OBSERVACION])
        for fila in registros
    )


def registrar_entrada(matricula):
    try:
        if hoja_registros is None:
            return None

        nombre = buscar_nombre(matricula)
        if nombre:
            hora, fecha = obtener_hora_internet()
            laptop_id = socket.gethostname()

            fila_nueva = [
                matricula,  # A Matrícula
                nombre,     # B Nombre
                fecha,      # C Fecha
                hora,       # D Hora_Ingreso
                "",         # E Hora_Salida
                laptop_id,  # F Laptop_ID
                "S/N"       # G Observación
            ]

            with bloqueo_local_registro_prestamo():
                registros = hoja_registros.get_all_values()
                if _existe_sesion_activa_para_matricula(registros, matricula):
                    return None

                try:
                    hoja_registros.append_row(fila_nueva)
                except Exception as error_append:
                    print(f"Respuesta ambigua al registrar entrada: {error_append}")
                    try:
                        registros = hoja_registros.get_all_values()
                        if _existe_entrada_del_mismo_intento(
                            registros, matricula, nombre, fecha, hora, laptop_id
                        ):
                            return nombre
                    except Exception as error_verificacion:
                        print(f"No se pudo verificar la entrada ambigua: {error_verificacion}")
                    raise

            time.sleep(1)  # JUSTO AQUÍ (MUY IMPORTANTE)

            return nombre

    except Exception as e:
        print(f"Error al registrar entrada: {e}")

    return None


def _existe_entrada_del_mismo_intento(
    registros, matricula, nombre, fecha, hora_ingreso, laptop_id
):
    for fila in reversed(registros):
        if (
            len(fila) > COL_OBSERVACION
            and fila[COL_MATRICULA] == matricula
            and fila[COL_NOMBRE] == nombre
            and fila[COL_FECHA] == fecha
            and fila[COL_HORA_INGRESO] == hora_ingreso
            and fila[COL_LAPTOP_ID] == laptop_id
            and fila[COL_HORA_SALIDA].strip() == ""
            and observacion_es_normal(fila[COL_OBSERVACION])
        ):
            return True
    return False


def registrar_salida_con_reintentos(
    nombre, matricula, max_reintentos=3, callback_intento=None
):
    laptop_actual = socket.gethostname()
    salida_iniciada = False
    pausas = (0, 1, 2)

    for numero_intento in range(1, max_reintentos + 1):
        if numero_intento > 1:
            time.sleep(pausas[min(numero_intento - 1, len(pausas) - 1)])
        if callback_intento:
            callback_intento(numero_intento, max_reintentos)

        try:
            if hoja_registros is None or not verificar_conexion_base_datos():
                raise ErrorConsultaGoogle("Google Sheets no responde")

            registros = hoja_registros.get_all_values()
            if not registros:
                raise ErrorConsultaGoogle("La hoja Registros está vacía")

            hora, _ = obtener_hora_internet()
            for i in reversed(range(len(registros))):
                fila = registros[i]
                if len(fila) <= COL_OBSERVACION or fila[COL_MATRICULA] != matricula:
                    continue

                hora_salida = fila[COL_HORA_SALIDA].strip()
                laptop_registro = fila[COL_LAPTOP_ID]
                observacion = fila[COL_OBSERVACION].strip()
                if not observacion_es_normal(observacion):
                    continue

                if laptop_registro == laptop_actual and hora_salida == "":
                    numero_fila = i + 1
                    salida_iniciada = True
                    try:
                        hoja_registros.update_cell(
                            numero_fila, COL_HORA_SALIDA + 1, hora
                        )
                    except Exception as error_update:
                        print(f"Respuesta ambigua al registrar salida: {error_update}")

                    salida_guardada = hoja_registros.cell(
                        numero_fila, COL_HORA_SALIDA + 1
                    ).value
                    if (salida_guardada or "").strip() != "":
                        registrar_evento_tecnico(
                            f"SHEETS_SALIDA intento={numero_intento} resultado=OK"
                        )
                        return True
                    raise ErrorConsultaGoogle("No se confirmó Hora_Salida")

                if hora_salida == "" and laptop_registro != laptop_actual:
                    registrar_evento_tecnico(
                        f"SHEETS_SALIDA intento={numero_intento} resultado=ERROR"
                    )
                    return False

                if laptop_registro == laptop_actual and salida_iniciada:
                    registrar_evento_tecnico(
                        f"SHEETS_SALIDA intento={numero_intento} resultado=OK"
                    )
                    return True
                registrar_evento_tecnico(
                    f"SHEETS_SALIDA intento={numero_intento} resultado=ERROR"
                )
                return False

            raise ErrorConsultaGoogle("No se encontró una sesión activa")
        except Exception as error:
            print(f"Error intento {numero_intento}: {error}")
            registrar_evento_tecnico(
                f"SHEETS_SALIDA intento={numero_intento} resultado=ERROR"
            )

    return False


def mostrar_ventana_espera_registro(ventana_entrega, matricula, nombre,
                                    modal_progreso):
    """Registra la salida actualizando el modal abierto de devolución."""
    ventana_espera, titulo_label, mensaje_label, estado_label = modal_progreso

    def actualizar_progreso(estado):
        estado_label.config(text=estado)

    def intentar_registro():
        def notificar_intento(intento, total):
            ejecutar_en_ui(
                actualizar_progreso,
                "Guardando devolución..."
            )

        exito = registrar_salida_con_reintentos(
            nombre, matricula, max_reintentos=3,
            callback_intento=notificar_intento
        )
        if exito:
            ejecutar_en_ui(
                actualizar_progreso,
                "Registro completado. Finalizando..."
            )
            time.sleep(1)
        return exito

    def finalizar_registro(exito):
        cerrar_modal_verificacion_ubicacion(ventana_espera)
        if exito:
            ventana_entrega.destroy()
            apagar_windows()
        else:
            mostrar_error(
                "Error",
                "No se pudo registrar la salida.\n\n"
                "La laptop no se apagará.\n"
                "Por favor intenta nuevamente."
            )

    # Ejecutar en un hilo separado
    def ejecutar_registro():
        exito = intentar_registro()
        ejecutar_en_ui(finalizar_registro, exito)

    threading.Thread(target=ejecutar_registro, daemon=True).start()

    return ventana_espera

def entregar_y_apagar(ventana, matricula, nombre, boton=None):
    print("[DEVOLUCION] entrando a entregar_y_apagar")

    if boton is None:
        print("[DEVOLUCION] cancelada: botón no disponible")
        return

    if boton:
        boton.config(
            state="disabled",
            text=(
                "Verificando ubicación..."
                if PROXIMIDAD_HABILITADA else "Verificando conexión..."
            )
        )
        print("[DEVOLUCION] botón deshabilitado")
        if PROXIMIDAD_HABILITADA:
            modal_progreso = crear_modal_verificacion_ubicacion(
                ventana,
                "Asegúrate de estar cerca del carrito para continuar con la devolución."
            )
        else:
            modal_progreso = crear_modal_verificacion_ubicacion(
                ventana,
                "Estamos verificando la información y guardando la entrega del equipo. "
                "Por favor, espera.",
                titulo="Registrando devolución",
                estado="Procesando..."
            )
        print("[DEVOLUCION] modal creado")
        ventana_ubicacion, titulo_ubicacion, mensaje_ubicacion, estado_ubicacion = modal_progreso
        proximidad_visual_activa = {"valor": True}

        def actualizar_intento_ubicacion(intento, total):
            if proximidad_visual_activa["valor"]:
                try:
                    estado_ubicacion.config(text=f"Intentando {intento} de {total}...")
                except tk.TclError:
                    pass

        def notificar_intento_ubicacion(intento, total):
            ejecutar_en_ui(actualizar_intento_ubicacion, intento, total)

        def comprobar_proximidad():
            print("[DEVOLUCION] worker de proximidad ejecutándose")
            try:
                proximidad_valida, estado_proximidad = verificar_proximidad_carrito(
                    callback_intento=notificar_intento_ubicacion
                )
            except Exception:
                print("[DEVOLUCION] excepción en proximidad")
                traceback.print_exc()
                proximidad_valida = False
                estado_proximidad = ESTADO_ERROR

            ejecutar_en_ui(
                procesar_resultado_proximidad,
                proximidad_valida,
                estado_proximidad
            )

        def procesar_resultado_proximidad(proximidad_valida, estado_proximidad):
            nonlocal modal_progreso, ventana_ubicacion
            nonlocal titulo_ubicacion, mensaje_ubicacion, estado_ubicacion
            proximidad_visual_activa["valor"] = False

            if not proximidad_valida:
                permite_contingencia = estado_proximidad in (
                    ESTADO_NO_DETECTADO,
                    ESTADO_ERROR
                )
                if permite_contingencia:
                    iniciar_contexto_contingencia(
                        estado_proximidad,
                        "DEVOLUCION",
                        matricula,
                        ventana
                    )
                else:
                    cerrar_modal_verificacion_ubicacion(ventana_ubicacion)

                mostrar_fallo_proximidad(
                    estado_proximidad,
                    parent=ventana,
                    al_crear=(
                        registrar_dialogo_fallo_contingencia
                        if permite_contingencia else None
                    )
                )

                if permite_contingencia and finalizar_contexto_contingencia():
                    cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                    modal_progreso = crear_modal_verificacion_ubicacion(
                        ventana,
                        "Estamos verificando la información y guardando la entrega del equipo. "
                        "Por favor, espera.",
                        titulo="Registrando devolución",
                        estado="Procesando..."
                    )
                    (
                        ventana_ubicacion,
                        titulo_ubicacion,
                        mensaje_ubicacion,
                        estado_ubicacion
                    ) = modal_progreso
                    procesar_resultado_proximidad(True, None)
                    return

                if permite_contingencia:
                    cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                if boton:
                    boton.config(state="normal", text="ENTREGAR Y APAGAR")
                return

            if boton:
                boton.config(text="Verificando conexión...")
            actualizar_modal_progreso(
                ventana_ubicacion,
                titulo_ubicacion,
                mensaje_ubicacion,
                estado_ubicacion,
                "Registrando devolución",
                "Estamos verificando la información y guardando la entrega del equipo. "
                "Por favor, espera.",
                "Procesando..."
            )
            threading.Thread(target=comprobar_devolucion, daemon=True).start()

        def comprobar_devolucion():
            def notificar_intento(intento, total):
                ejecutar_en_ui(actualizar_intento_conexion, intento, total)

            if not verificar_sheets_con_reintentos(
                max_intentos=3, callback_intento=notificar_intento
            ):
                ejecutar_en_ui(procesar_resultado_devolucion, "error_conexion", None)
                return

            try:
                registros = hoja_registros.get_all_values()
                ejecutar_en_ui(procesar_resultado_devolucion, "ok", registros)
            except Exception:
                ejecutar_en_ui(procesar_resultado_devolucion, "error_conexion", None)

        def actualizar_intento_conexion(intento, total):
            estado_ubicacion.config(
                text="Verificando conexión..."
            )

        def procesar_resultado_devolucion(estado, registros):
            if estado != "ok":
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                if boton:
                    boton.config(state="normal", text="ENTREGAR Y APAGAR")
                if estado == "sin_internet":
                    mostrar_error(
                        "Sin acceso a Internet",
                        "No hay conexión.\n\nLa laptop NO se apagará.",
                        parent=ventana
                    )
                else:
                    mostrar_error(
                        "Error de conexión",
                        "No fue posible verificar la conexión.\n\nIntenta nuevamente.",
                        parent=ventana
                    )
                return

            laptop_actual = socket.gethostname()
            fila_encontrada = None
            for fila in reversed(registros):
                if (
                    len(fila) > COL_OBSERVACION
                    and fila[COL_MATRICULA] == matricula
                    and fila[COL_LAPTOP_ID] == laptop_actual
                    and fila[COL_HORA_SALIDA].strip() == ""
                    and observacion_es_normal(fila[COL_OBSERVACION])
                ):
                    fila_encontrada = fila
                    break

            # Conserva los avisos existentes para una sesión ya cerrada, pero una
            # no entrega contabilizada nunca representa la devolución actual.
            if not fila_encontrada:
                for fila in reversed(registros):
                    if (
                        len(fila) > COL_OBSERVACION
                        and fila[COL_MATRICULA] == matricula
                        and fila[COL_LAPTOP_ID] == laptop_actual
                        and fila[COL_OBSERVACION].strip() != "NO_ENTREGA_CONTADA"
                    ):
                        fila_encontrada = fila
                        break

            if not fila_encontrada:
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                if boton:
                    boton.config(state="normal", text="ENTREGAR Y APAGAR")
                mostrar_error(
                    "Error",
                    "No se encontró una sesión en esta laptop.",
                    parent=ventana
                )
                return

            hora_salida = fila_encontrada[COL_HORA_SALIDA].strip()
            observacion = fila_encontrada[COL_OBSERVACION].strip()

            if observacion == "CIERRE_AUTOMATICO_POR_NUEVA_SESION":
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                mostrar_advertencia(
                    "Sesión cerrada automáticamente",
                    "Esta sesión ya fue cerrada automáticamente porque se inició en otro equipo.\n\n"
                    "La laptop se apagará.",
                    parent=ventana
                )
                ventana.destroy()
                apagar_windows()
                return

            if hora_salida != "":
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                mostrar_informacion(
                    "Entrega ya registrada",
                    "La salida ya fue registrada.\n\nLa laptop se apagará.",
                    parent=ventana
                )
                ventana.destroy()
                apagar_windows()
                return

            mostrar_ventana_espera_registro(
                ventana, matricula, nombre, modal_progreso
            )

    if PROXIMIDAD_HABILITADA:
        print("[DEVOLUCION] iniciando proximidad")
        threading.Thread(target=comprobar_proximidad, daemon=True).start()
        print("[DEVOLUCION] worker iniciado")
    else:
        print("[DEVOLUCION] proximidad deshabilitada; verificando devolución")
        threading.Thread(target=comprobar_devolucion, daemon=True).start()
    return



def mostrar_ventana_entrega(nombre, matricula):
    global ventana_entrega_activa

    ventana_entrega = tk.Toplevel()
    ventana_entrega_activa = ventana_entrega

    def limpiar_referencia_entrega(event):
        global ventana_entrega_activa
        if event.widget == ventana_entrega:
            ventana_entrega_activa = None

    ventana_entrega.bind("<Destroy>", limpiar_referencia_entrega)

    
    # CLAVE: ventana independiente
    ventana_entrega.transient(None)
    ventana_entrega.title("Entrega de Laptop")
    ventana_entrega.resizable(False, False)
    ventana_entrega.configure(bg=COLOR_FONDO)

    # ? Bloquear cerrar, pero NO minimizar
    ventana_entrega.protocol("WM_DELETE_WINDOW", lambda: None)

    # Permitir minimizar
    ventana_entrega.minsize(420, 480)

    frame_principal = tk.Frame(
        ventana_entrega,
        bg=COLOR_TARJETA,
        padx=35,
        pady=30
    )
    frame_principal.pack(fill=tk.BOTH, expand=True)

    if TEMA_INTERFAZ == "moderno":
        tk.Frame(frame_principal, height=4, bg=COLOR_EXITO).pack(
            fill=tk.X, pady=(0, 18)
        )

    # Logo
    logo_entrega = cargar_logo(RUTA_UTP, 120, 70)
    if logo_entrega:
        lbl_logo = tk.Label(frame_principal, image=logo_entrega, bg=COLOR_TARJETA)
        lbl_logo.image = logo_entrega
        lbl_logo.pack(pady=(0, 20))


    tk.Label(
        frame_principal,
        text="Sistema de Control de Laptops",
        font=FUENTE_SUBTITULO,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 5))

    tk.Label(
        frame_principal,
        text=f"Bienvenid@,\n{nombre}",
        font=("Segoe UI", 16, "bold"),
        fg=COLOR_EXITO,
        bg=COLOR_TARJETA,
        justify="center",
        wraplength=340
    ).pack(pady=(0, 25))


    # Botón principal
    def iniciar_devolucion_desde_boton():
        print("[DEVOLUCION] clic ENTREGAR Y APAGAR")
        try:
            entregar_y_apagar(
                ventana_entrega,
                matricula,
                nombre,
                boton=btn_entregar
            )
        except Exception:
            print("[DEVOLUCION] excepción al iniciar la devolución")
            traceback.print_exc()
            try:
                if btn_entregar.winfo_exists():
                    btn_entregar.config(
                        state="normal",
                        text="ENTREGAR Y APAGAR"
                    )
            except tk.TclError:
                pass

    btn_entregar = tk.Button(
        frame_principal,
        text="ENTREGAR Y APAGAR",
        padx=30,
        pady=15,
        command=iniciar_devolucion_desde_boton,
        **estilo_boton("error")
    )
    btn_entregar.pack(pady=20)

    tk.Label(
        frame_principal,
        text="La laptop se apagará automáticamente",
        font=FUENTE_PEQ,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(pady=(10, 0))



    # Centrar ventana
    ancho, alto = 420, 480
    x = (ventana_entrega.winfo_screenwidth() - ancho) // 2
    y = (ventana_entrega.winfo_screenheight() - alto) // 2
    ventana_entrega.geometry(f"{ancho}x{alto}+{x}+{y}")
    ventana_entrega.deiconify()
    ventana_entrega.lift()
    ventana_entrega.focus_force()

def formatear_fecha(fecha_entrada):
    try:
        fecha_obj = datetime.strptime(fecha_entrada, "%Y-%m-%d")
        meses_espanol = {
            1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
            5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
            9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
        }
        dia = fecha_obj.day
        mes = meses_espanol[fecha_obj.month]
        return f"{dia} de {mes}"
    except:
        return fecha_entrada

def mostrar_aviso_entrega_pendiente(fecha_entrada, laptop_id):
    fecha_formateada = formatear_fecha(fecha_entrada)
    
    mensaje = f"Olvidaste entregar la laptop\n\n"
    mensaje += f"Laptop: {laptop_id}\n"
    mensaje += f"Fecha: {fecha_formateada}\n\n"
    mensaje += "Recuerda siempre usar el botón 'Entregar y Apagar'"
    
    mostrar_advertencia("Entrega Pendiente", mensaje)
def verificar_sesion_activa_en_otra_laptop(matricula):
    """
    Verifica si la matrícula tiene una sesión activa
    en OTRA laptop diferente a la actual.
    """
    try:
        if hoja_registros is None:
            raise ErrorConsultaGoogle("La hoja Registros no está disponible")

        laptop_actual = socket.gethostname()
        registros = hoja_registros.get_all_values()

        for fila in reversed(registros):
            if len(fila) > COL_OBSERVACION and fila[0] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                laptop_registro = fila[COL_LAPTOP_ID]
                observacion = fila[COL_OBSERVACION].strip()


                if (
                    hora_salida == ""
                    and observacion_es_normal(observacion)
                    and laptop_registro != laptop_actual
                ):
                    return True, laptop_registro

    except ErrorConsultaGoogle:
        raise
    except Exception as e:
        print(f"Error al verificar sesión en otra laptop: {e}")
        raise ErrorConsultaGoogle("No se pudo verificar la sesión activa") from e

    return False, None


def mostrar_confirmacion_simple(nombre, matricula, identidad_confirmada=False):
    global entrada, ventana, procesando_sesion

    if not identidad_confirmada:
        activar_timeout_inactividad()
        respuesta = mostrar_confirmacion_personalizada(
            "Confirmación",
            f"¿Eres {nombre}?"
        )
        cancelar_timeout_inactividad()

        if not respuesta:
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return False

    cancelar_timeout_inactividad()
    cambiar_estado("Consultando sesión...", COLOR_ADVERTENCIA)

    def consultar_sesion():
        try:
            resultado = verificar_sesion_activa_en_otra_laptop(matricula)
            ejecutar_en_ui(procesar_sesion, "ok", *resultado)
        except ErrorConsultaGoogle:
            ejecutar_en_ui(procesar_sesion, "error", False, None)

    def procesar_sesion(resultado_consulta, sesion_activa, laptop_otro):
        if resultado_consulta == "error":
            mostrar_error(
                "No se pudo verificar la información",
                "No fue posible verificar la conexión. Intenta nuevamente."
            )
            reiniciar_estado_sistema()
            return

        if sesion_activa:
            activar_timeout_inactividad()
            continuar = mostrar_confirmacion_personalizada(
                "Sesión activa detectada",
                "Se detectó una sesión activa en otro equipo.\n\n"
                f"Laptop anterior: {laptop_otro}\n\n"
                "Si continúas:\n"
                "• La sesión anterior se cerrará automáticamente\n"
                "• Se registrará como NO ENTREGA\n"
                "• Esta acción quedará registrada\n\n"
                "¿Deseas continuar en este equipo?"
            )
            cancelar_timeout_inactividad()
            if not continuar:
                reiniciar_estado_sistema()
                entrada.delete(0, tk.END)
                entrada.focus()
                return

            cambiar_estado("Actualizando sesión anterior...", COLOR_ADVERTENCIA)
            threading.Thread(target=cerrar_sesion_anterior, daemon=True).start()
            return

        iniciar_validacion_identidad()

    def cerrar_sesion_anterior():
        try:
            exito = cerrar_sesion_anterior_y_contar_no_entrega(matricula)
        except ErrorConsultaGoogle:
            exito = False
        ejecutar_en_ui(procesar_cierre_anterior, exito)

    def procesar_cierre_anterior(exito):
        if not exito:
            mostrar_error(
                "Error",
                "No se pudo cerrar la sesión anterior.\n\nIntenta nuevamente."
            )
            reiniciar_estado_sistema()
            return
        iniciar_validacion_identidad()

    def iniciar_validacion_identidad():
        cambiar_estado("Validando identidad...", COLOR_ADVERTENCIA)
        threading.Thread(target=consultar_identidad, daemon=True).start()

    def consultar_identidad():
        curp_real = buscar_curp(matricula)
        ejecutar_en_ui(procesar_identidad, curp_real)

    def procesar_identidad(curp_real):
        def cancelar_validacion_curp():
            cancelar_timeout_inactividad()
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus_set()

        if not curp_real:
            mostrar_error(
                "Error",
                "No se pudo validar la identidad.\nContacta al administrador."
            )
            cancelar_validacion_curp()
            return

        for intento in range(1, 4):
            curp_ingresada = pedir_curp_ultimos_2(ventana)
            if not curp_ingresada:
                cancelar_validacion_curp()
                return

            if validar_curp_ultimos_2(curp_real, curp_ingresada):
                break

            intentos_restantes = 3 - intento
            if intentos_restantes:
                texto_intentos = (
                    "Te queda 1 intento."
                    if intentos_restantes == 1
                    else f"Te quedan {intentos_restantes} intentos."
                )
                mostrar_error(
                    "Datos incorrectos",
                    "Los datos ingresados no coinciden.",
                    texto_destacado=texto_intentos,
                    texto_boton="INTENTAR NUEVAMENTE"
                )
                if flujo_cancelado_por_inactividad:
                    cancelar_validacion_curp()
                    return
                continue

            mostrar_error(
                "Acceso denegado",
                "Los datos no coinciden.\n\n"
                "Se agotaron los intentos de validación. Sesión cancelada."
            )
            cancelar_validacion_curp()
            return

        cancelar_timeout_inactividad()
        cambiar_estado("Consultando registros...", COLOR_ADVERTENCIA)
        threading.Thread(target=consultar_control, daemon=True).start()

    def consultar_control():
        try:
            procesar_no_entrega_si_corresponde(matricula)
            no_entregas, estado = obtener_control_alumno(matricula)
            tiene_pendiente, fecha_entrada, laptop_id = verificar_entrega_pendiente(matricula)
            ejecutar_en_ui(
                procesar_control,
                no_entregas,
                estado,
                tiene_pendiente,
                fecha_entrada,
                laptop_id
            )
        except ErrorConsultaGoogle:
            ejecutar_en_ui(procesar_error_control)

    def procesar_error_control():
        mostrar_error(
            "No se pudo verificar la información",
            "No fue posible verificar la conexión. Intenta nuevamente."
        )
        reiniciar_estado_sistema()

    def procesar_control(no_entregas, estado, tiene_pendiente, fecha_entrada, laptop_id):
        estado = normalizar_estado_alumno(estado)
        if estado is None:
            registrar_evento_tecnico(
                "ESTADO_ALUMNO_BLOQUEADO valor=DESCONOCIDO"
            )
            mostrar_error(
                "Acceso denegado",
                "El estado del alumno no permite registrar un préstamo.\n\n"
                "Contacta al administrador."
            )
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return

        if no_entregas >= 2:
            activar_timeout_inactividad()
            mostrar_ventana_control_unificada(
                matricula=matricula,
                nombre=nombre,
                tiene_pendiente=tiene_pendiente,
                fecha_entrada=fecha_entrada,
                laptop_id=laptop_id,
                no_entregas=no_entregas,
                estado=("" if estado == ESTADO_VACIO_VALIDO else estado)
            )
            cancelar_timeout_inactividad()
            if flujo_cancelado_por_inactividad:
                reiniciar_estado_sistema()
                return
            if estado == "SANCIONADO":
                reiniciar_estado_sistema()
                entrada.delete(0, tk.END)
                entrada.focus()
                return

        if not estado_alumno_permite_prestamo(estado):
            registrar_evento_tecnico(
                f"ESTADO_ALUMNO_BLOQUEADO valor={estado}"
            )
            mostrar_error(
                "Acceso denegado",
                "El estado del alumno no permite registrar un préstamo.\n\n"
                "Contacta al administrador."
            )
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return

        cambiar_estado("Registrando préstamo...", COLOR_ADVERTENCIA)
        threading.Thread(target=registrar_prestamo, daemon=True).start()

    def registrar_prestamo():
        resultado = registrar_entrada(matricula)
        ejecutar_en_ui(finalizar_prestamo, resultado)

    def finalizar_prestamo(resultado):
        global procesando_sesion
        if not resultado:
            mostrar_error("Error", "No se pudo registrar la entrada.")
            reiniciar_estado_sistema()
            return

        procesando_sesion = False
        ventana.withdraw()
        ventana.after(200, lambda: mostrar_ventana_entrega(nombre, matricula))

    threading.Thread(target=consultar_sesion, daemon=True).start()
    return True


def cambiar_estado(texto, color=None):
    if threading.current_thread() is not threading.main_thread():
        ejecutar_en_ui(cambiar_estado, texto, color)
        return

    estado_var.set(texto)
    if color:
        estado_label.config(fg=color)
    ventana.update_idletasks()

def reiniciar_estado_sistema():
    """Reinicia el estado del sistema después de una operación"""
    global procesando_sesion
    procesando_sesion = False
    btn_entrar.config(state="normal")
    # Actualizar estado de conexión actual
    if verificar_conexion_base_datos():
        cambiar_estado("Conectado", COLOR_EXITO)
    else:
        cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
        if acepta_estado_equipo:
            acepta_estado_equipo.set(False)
            try:
                pass
            except:
                pass



def mostrar_aviso_internet_bloqueante():
    """Muestra un aviso emergente bloqueante que no se puede mover, cerrar ni minimizar"""
    global aviso_internet

    try:
        if aviso_internet and aviso_internet.winfo_exists():
            aviso_internet.lift()
            return
    except:
        aviso_internet = None
    aviso_internet = tk.Toplevel(ventana)
    aviso_internet.title("Conexión Requerida")
    
    # Hacer la ventana completamente bloqueante
    aviso_internet.attributes('-topmost', True)
    aviso_internet.resizable(False, False)
    aviso_internet.overrideredirect(True)
    
    # Configurar tamaño y posición
    ancho_aviso = 500
    alto_aviso = 320
    
    pantalla_ancho = aviso_internet.winfo_screenwidth()
    pantalla_alto = aviso_internet.winfo_screenheight()
    x = (pantalla_ancho - ancho_aviso) // 2
    y = (pantalla_alto - alto_aviso) // 2
    
    aviso_internet.geometry(f"{ancho_aviso}x{alto_aviso}+{x}+{y}")
    aviso_internet.configure(bg=COLOR_PRIMARIO)
    
    # Frame principal
    frame_aviso = tk.Frame(aviso_internet, bg=COLOR_TARJETA, bd=0, relief=tk.FLAT)
    frame_aviso.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
    
    # Logo
    logo_aviso = cargar_logo(RUTA_UTP, 70, 45)
    if logo_aviso:
        lbl_logo_aviso = tk.Label(frame_aviso, image=logo_aviso, bg=COLOR_TARJETA)
        lbl_logo_aviso.image = logo_aviso
        lbl_logo_aviso.pack(pady=(25, 10))
    
    # Icono de advertencia
    tk.Label(frame_aviso, 
             text="!",
             font=("Segoe UI", 32),
             bg=COLOR_TARJETA).pack(pady=(10, 15))
    
    # Mensaje principal
    tk.Label(frame_aviso, 
             text="NO SE HA DETECTADO CONEXIÓN A INTERNET",
             font=("Segoe UI", 13, "bold"),
             fg=COLOR_ERROR,
             bg=COLOR_TARJETA).pack(pady=(0, 10))
    
    # Instrucciones
    tk.Label(frame_aviso, 
             text="Para utilizar el sistema es necesario contar con conexión a internet.",
             font=FUENTE_CUERPO,
             fg=COLOR_TEXTO,
             bg=COLOR_TARJETA,
             wraplength=400,
             justify=tk.CENTER).pack(pady=(0, 5))
    
    tk.Label(frame_aviso, 
             text="Conecte la computadora a internet para continuar.",
             font=FUENTE_CUERPO_BOLD,
             fg=COLOR_TEXTO,
             bg=COLOR_TARJETA,
             wraplength=400,
             justify=tk.CENTER).pack(pady=(0, 20))
    
    # Barra de progreso animada
    progress_frame = tk.Frame(frame_aviso, bg=COLOR_TARJETA)
    progress_frame.pack(pady=15)
    
    progress_bar = tk.Frame(progress_frame, height=8, bg=COLOR_BORDE, width=350)
    progress_bar.pack()
    progress_bar_inner = tk.Frame(progress_bar, height=8, bg=COLOR_PRIMARIO, width=0)
    progress_bar_inner.place(relx=0, rely=0, relheight=1)
    
    # Mensaje de estado
    estado_label = tk.Label(frame_aviso, 
                           text="Verificando conexión...",
                           font=FUENTE_PEQ,
                           fg=COLOR_TEXTO_SECUNDARIO,
                           bg=COLOR_TARJETA)
    estado_label.pack(pady=5)
    
    # Contador de tiempo
    tiempo_inicio = time.time()
    tiempo_label = tk.Label(frame_aviso, 
                           text="Tiempo esperando: 0 segundos",
                           font=FUENTE_PEQ,
                           fg=COLOR_TEXTO_SECUNDARIO,
                           bg=COLOR_TARJETA)
    tiempo_label.pack(pady=2)
    
    # Función para animar la barra de progreso
    def animar_barra_progreso():
        nonlocal progreso_actual, direccion
        if direccion == 1:
            progreso_actual += 5
            if progreso_actual >= 350:
                progreso_actual = 350
                direccion = -1
        else:
            progreso_actual -= 5
            if progreso_actual <= 0:
                progreso_actual = 0
                direccion = 1
        
        try:
            if not aviso_internet or not aviso_internet.winfo_exists():
                return
            progress_bar_inner.config(width=progreso_actual)
            aviso_internet.after(50, animar_barra_progreso)
        except:
            return
    
    # Función para verificar internet continuamente sin congelar la ventana
    def verificar_internet_continuamente():
        global verificacion_aviso_en_curso

        if verificacion_aviso_en_curso:
            return

        try:
            if not aviso_internet or not aviso_internet.winfo_exists():
                return
        except:
            return

        tiempo_transcurrido = int(time.time() - tiempo_inicio)
        tiempo_label.config(text=f"Tiempo esperando: {tiempo_transcurrido} segundos")
        estado_label.config(text=f"Verificando conexión... ({time.strftime('%H:%M:%S')})")
        verificacion_aviso_en_curso = True

        def revisar_conexion():
            conectado = verificar_internet()
            if conectado:
                conectado = conectar_google_sheets(internet_verificado=True)
            ejecutar_en_ui(procesar_resultado_conexion, conectado)

        threading.Thread(target=revisar_conexion, daemon=True).start()

    def procesar_resultado_conexion(conectado):
        global verificacion_aviso_en_curso
        verificacion_aviso_en_curso = False

        try:
            if not aviso_internet or not aviso_internet.winfo_exists():
                return
        except:
            return

        if conectado:
            cerrar_aviso_internet()
            cambiar_estado("Conectado", COLOR_EXITO)
        else:
            aviso_internet.after(1000, verificar_internet_continuamente)
    
    # Iniciar animación y verificación
    progreso_actual = 0
    direccion = 1
    animar_barra_progreso()
    aviso_internet.after(500, verificar_internet_continuamente)
def iniciar_sesion():
    global procesando_sesion, flujo_cancelado_por_inactividad
    cancelar_timeout_inactividad()

    if procesando_sesion:
        return


    flujo_cancelado_por_inactividad = False
    procesando_sesion = True
    btn_entrar.config(state="disabled")

    matricula = entrada.get().strip()

    if not matricula:
        mostrar_advertencia("Campo vacío", "Por favor ingresa tu matrícula.")
        entrada.focus()
        reiniciar_estado_sistema()
        return
        # Validar aceptación de estado del equipo
    if not acepta_estado_equipo.get():
        mostrar_advertencia(
            "Confirmación requerida",
            "Debes confirmar el estado de la laptop para continuar."
        )
        reiniciar_estado_sistema()
        return

    cancelar_timeout_inactividad()
    ventana_ubicacion = None
    titulo_ubicacion = None
    mensaje_ubicacion = None
    estado_ubicacion = None
    proximidad_visual_activa = {"valor": False}
    alumno_confirmado = {"nombre": None}


    def consultar_acceso():
        def notificar_intento(intento, total):
            ejecutar_en_ui(
                cambiar_estado,
                "Verificando conexión...",
                COLOR_ADVERTENCIA
            )

        if not verificar_sheets_con_reintentos(
            max_intentos=3, callback_intento=notificar_intento
        ):
            ejecutar_en_ui(procesar_acceso_seguro, "error_conexion", None)
            return

        nombre = buscar_nombre(matricula)
        ejecutar_en_ui(procesar_acceso_seguro, "ok", nombre)

    def procesar_acceso_seguro(estado, nombre):
        try:
            procesar_acceso(estado, nombre)
        except Exception:
            traceback.print_exc()
            reiniciar_estado_sistema()

    def procesar_acceso(estado, nombre):
        nonlocal ventana_ubicacion, titulo_ubicacion
        nonlocal mensaje_ubicacion, estado_ubicacion
        if estado == "sin_internet":
            mostrar_error(
                "Sin acceso a Internet",
                "No hay conexión disponible.\n\nConecte la computadora e intente nuevamente."
            )
            reiniciar_estado_sistema()
            return

        if estado == "error_conexion":
            mostrar_error(
                "Error de conexión",
                "No fue posible verificar la conexión.\n\nIntenta nuevamente."
            )
            reiniciar_estado_sistema()
            return

        if not nombre:
            mostrar_error(
                "Matrícula no válida",
                "La matrícula no está registrada.\n\nVerifica o contacta al administrador."
            )
            entrada.delete(0, tk.END)
            entrada.focus()
            reiniciar_estado_sistema()
            return

        activar_timeout_inactividad()
        confirmado = mostrar_confirmacion_personalizada(
            "Confirmación",
            f"¿Eres {nombre}?"
        )
        cancelar_timeout_inactividad()
        if not confirmado:
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return

        alumno_confirmado["nombre"] = nombre
        if not PROXIMIDAD_HABILITADA:
            mostrar_confirmacion_simple(
                alumno_confirmado["nombre"],
                matricula,
                identidad_confirmada=True
            )
            return

        (
            ventana_ubicacion,
            titulo_ubicacion,
            mensaje_ubicacion,
            estado_ubicacion
        ) = crear_modal_verificacion_ubicacion(
            ventana,
            "Asegúrate de estar cerca del carrito para continuar."
        )
        proximidad_visual_activa["valor"] = True
        threading.Thread(target=comprobar_proximidad, daemon=True).start()

    def comprobar_proximidad():
        try:
            proximidad_valida, estado_proximidad = verificar_proximidad_carrito(
                callback_intento=notificar_intento_ubicacion
            )
        except Exception:
            proximidad_valida = False
            estado_proximidad = ESTADO_ERROR

        ejecutar_en_ui(
            procesar_proximidad,
            proximidad_valida,
            estado_proximidad
        )

    def actualizar_intento_ubicacion(intento, total):
        if proximidad_visual_activa["valor"]:
            try:
                estado_ubicacion.config(text=f"Intentando {intento} de {total}...")
            except tk.TclError:
                pass

    def notificar_intento_ubicacion(intento, total):
        ejecutar_en_ui(actualizar_intento_ubicacion, intento, total)

    def procesar_proximidad(proximidad_valida, estado_proximidad):
        proximidad_visual_activa["valor"] = False

        if not proximidad_valida:
            permite_contingencia = estado_proximidad in (
                ESTADO_NO_DETECTADO,
                ESTADO_ERROR
            )
            if permite_contingencia:
                iniciar_contexto_contingencia(
                    estado_proximidad,
                    "PRESTAMO",
                    matricula,
                    ventana
                )
            else:
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)

            try:
                activar_timeout_inactividad()
                mostrar_fallo_proximidad(
                    estado_proximidad,
                    parent=ventana,
                    al_crear=(
                        registrar_dialogo_fallo_contingencia
                        if permite_contingencia else None
                    )
                )
            except Exception:
                traceback.print_exc()
            finally:
                cancelar_timeout_inactividad()

            if permite_contingencia and finalizar_contexto_contingencia():
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
                procesar_proximidad(True, None)
                return

            if permite_contingencia:
                cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
            reiniciar_estado_sistema()
            return

        cerrar_modal_verificacion_ubicacion(ventana_ubicacion)
        mostrar_confirmacion_simple(
            alumno_confirmado["nombre"],
            matricula,
            identidad_confirmada=True
        )

    cambiar_estado("Verificando conexión...", COLOR_ADVERTENCIA)
    threading.Thread(target=consultar_acceso, daemon=True).start()




def crear_pantalla_login():
    global entrada, btn_entrar, estado_var, estado_label, chk_label

    for widget in ventana.winfo_children():
        widget.destroy()

    ventana.configure(bg=COLOR_FONDO)

    main_container = tk.Frame(ventana, bg=COLOR_FONDO)
    main_container.pack(fill=tk.BOTH, expand=True)

    fondo_marca_agua = crear_fondo_marca_agua_login(
        RUTA_UTP,
        ventana.winfo_screenwidth(),
        ventana.winfo_screenheight()
    )
    if fondo_marca_agua:
        marca_agua_fondo = tk.Label(
            main_container,
            image=fondo_marca_agua,
            bg=COLOR_FONDO,
            bd=0
        )
        marca_agua_fondo.image = fondo_marca_agua
        marca_agua_fondo.place(x=0, y=0, relwidth=1, relheight=1)

    card = tk.Frame(
        main_container,
        bg=COLOR_TARJETA,
        width=TAMANOS["tarjeta_login_ancho"],
        height=TAMANOS["tarjeta_login_alto"],
        highlightthickness=1,
        highlightbackground=COLOR_BORDE
    )
    card.place(relx=0.5, rely=0.5, anchor="center")
    card.pack_propagate(False)

    if TEMA_INTERFAZ == "moderno":
        tk.Frame(card, height=5, bg=COLOR_PRIMARIO).pack(fill=tk.X)

    # LOGO
    logo = cargar_logo(RUTA_UTP, 150, 95)
    if logo:
        lbl_logo = tk.Label(card, image=logo, bg=COLOR_TARJETA)
        lbl_logo.image = logo
        lbl_logo.pack(pady=(35, 15))

    tk.Label(
        card,
        text="SISTEMA DE CONTROL\nDE LAPTOPS",
        font=("Segoe UI", 19 if TEMA_INTERFAZ == "moderno" else 18, "bold"),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA
    ).pack()

    tk.Label(
        card,
        text="Universidad Tecnológica de Parral",
        font=("Segoe UI", 11),
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(pady=(0, 30))

    # ENTRADA
    input_frame = tk.Frame(card, bg=COLOR_TARJETA)
    input_frame.pack(padx=50, fill=tk.X)

    tk.Label(
        input_frame,
        text="INGRESA TU MATRÍCULA",
        font=("Segoe UI", 10, "bold"),
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(anchor="w", pady=(0, 8))

    entrada = tk.Entry(
        input_frame,
        justify="center",
        **estilo_entrada()
    )
    entrada.pack(fill=tk.X, ipady=12, pady=(0, 18))

    def actualizar_timeout_desde_login(event=None):
        if entrada.get().strip() or acepta_estado_equipo.get():
            if timeout_inactividad_activo:
                reiniciar_timeout_inactividad()
            else:
                activar_timeout_inactividad()
        else:
            cancelar_timeout_inactividad()

    entrada.bind("<KeyRelease>", actualizar_timeout_desde_login, add="+")

    # FORZAR FOCO CORRECTAMENTE
    ventana.after(200, lambda: entrada.focus_set())

        # =========================
    # VALIDACIÓN SOLO NÚMEROS
    # =========================
    def solo_numeros(P):
        return P.isdigit() or P == ""

    vcmd = (ventana.register(solo_numeros), "%P")
    entrada.config(validate="key", validatecommand=vcmd)

        # =========================
        # CONFIRMACIÓN ESTADO LAPTOP
        # =========================
    global acepta_estado_equipo
    acepta_estado_equipo = tk.BooleanVar(value=False)

    check_frame = tk.Frame(card, bg=COLOR_TARJETA)
    check_frame.pack(padx=50, pady=(0, 25), fill=tk.X)

    check_content = tk.Frame(check_frame, bg=COLOR_TARJETA)
    check_content.pack(anchor="w")

    chk_label = tk.Canvas(
        check_content,
        width=32,
        height=32,
        bg=COLOR_TARJETA,
        highlightthickness=0,
        cursor="hand2"
    )
    chk_label.pack(side=tk.LEFT, padx=(0, 14))

    def dibujar_check():
        chk_label.delete("all")
        chk_label.create_rectangle(
            4, 4, 28, 28,
            outline=COLOR_PRIMARIO,
            width=3,
            fill=COLOR_TARJETA
        )
        if acepta_estado_equipo.get():
            chk_label.create_line(
                9, 17, 15, 23, 25, 10,
                fill=COLOR_EXITO,
                width=4,
                capstyle=tk.ROUND,
                joinstyle=tk.ROUND
            )

    chk_text = tk.Label(
        check_content,
        text="Confirmo que la laptop está en buen estado\n" +
             "o ya reporté cualquier anomalía.",
        font=("Segoe UI", 10),
        fg=COLOR_TEXTO,
        bg=COLOR_TARJETA,
        justify="left",
        cursor="hand2",
        wraplength=300
    )
    chk_text.pack(side=tk.LEFT)

    def toggle_check(event=None):
        acepta_estado_equipo.set(not acepta_estado_equipo.get())
        actualizar_timeout_desde_login()

    acepta_estado_equipo.trace_add("write", lambda *_: dibujar_check())
    dibujar_check()
    chk_label.bind("<Button-1>", toggle_check)
    chk_text.bind("<Button-1>", toggle_check)

        # Línea divisoria
    tk.Frame(check_frame, height=1, bg=COLOR_BORDE).pack(fill=tk.X, pady=(15, 0))

  
    entrada.bind("<Return>", lambda e: iniciar_sesion())


    # =========================
    # BOTÓN INGRESAR
    # =========================
    btn_entrar = tk.Button(
        card,
        text="INGRESAR AL SISTEMA",
        padx=20,
        pady=14,
        command=iniciar_sesion,
        **estilo_boton("principal")
    )
    btn_entrar.pack(fill=tk.X, padx=50, pady=(5, 20))

    # =========================
    # ESTADO DE CONEXIÓN
    # =========================
    estado_var = tk.StringVar(value="Iniciando...")
    estado_label = tk.Label(
        card,
        textvariable=estado_var,
        font=FUENTE_PEQ_BOLD,
        fg=COLOR_EXITO,
        bg=(COLORES["info_suave"] if TEMA_INTERFAZ == "moderno" else COLOR_TARJETA),
        padx=(12 if TEMA_INTERFAZ == "moderno" else 0),
        pady=(5 if TEMA_INTERFAZ == "moderno" else 0)
    )
    estado_label.pack(pady=(5, 10))

    # =========================
    # VERSIÓN (ABAJO DEL TODO)
    # =========================
    tk.Label(
        card,
        text=f"Versión v{VERSION_SISTEMA}",
        font=FUENTE_PEQ,
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(side=tk.BOTTOM, anchor="e", padx=12, pady=(5, 8))


def iniciar_conexion_en_segundo_plano():
    cambiar_estado("Verificando conexión...", COLOR_ADVERTENCIA)

    def revisar_y_conectar():
        if verificar_internet():
            conectar_google_sheets(internet_verificado=True)
        else:
            def mostrar_sin_conexion():
                cambiar_estado("Sin acceso a Internet", COLOR_ERROR)
                mostrar_aviso_internet_bloqueante()

            ejecutar_en_ui(mostrar_sin_conexion)

    threading.Thread(target=revisar_y_conectar, daemon=True).start()

def iniciar_verificacion_periodica():
    threading.Thread(target=verificar_conexion_periodicamente, daemon=True).start()


def notificar_interfaz_lista():
    """Confirma al bootstrap que la ventana principal ya procesa eventos."""
    ruta_senal = os.environ.get("REGISTRO_LAPTOP_APP_LISTA")
    token_senal = os.environ.get("REGISTRO_LAPTOP_APP_TOKEN")
    if not ruta_senal or not token_senal:
        return
    ruta_temporal = f"{ruta_senal}.{os.getpid()}.tmp"
    try:
        with open(ruta_temporal, "w", encoding="ascii") as archivo:
            archivo.write(token_senal)
        os.replace(ruta_temporal, ruta_senal)
        registrar_evento_tecnico("APP_LISTA enviada")
    except OSError as error:
        registrar_evento_tecnico(f"ERROR_APP_LISTA={type(error).__name__}: {error}")
        try:
            os.remove(ruta_temporal)
        except OSError:
            pass

# --- VENTANA PRINCIPAL ---
registrar_inicio_tecnico()
registrar_evento_tecnico(
    f"MODO_PRODUCCION={not MODO_PRUEBA} "
    f"PROXIMIDAD_HABILITADA={PROXIMIDAD_HABILITADA}"
)
ventana = tk.Tk()
ventana.bind_all("<Control-Alt-u>", manejar_atajo_administrativo)
ventana.bind_all("<KeyPress>", registrar_actividad_usuario, add="+")
ventana.bind_all("<ButtonPress>", registrar_actividad_usuario, add="+")
ventana.title(f"SISTEMA DE CONTROL DE LAPTOPS - UTP | v{VERSION_SISTEMA}")

# Pantalla completa sin bordes
ventana.attributes("-topmost", True)
ventana.state("zoomed")
ventana.overrideredirect(True)
ventana.configure(bg=COLOR_FONDO)

# AQUÍ SE CREA EL LOGIN
crear_pantalla_login()

# SEGURIDAD
ventana.protocol("WM_DELETE_WINDOW", cerrar_aplicacion)
ventana.bind_all("<Alt-F4>", bloquear_alt_f4)
ventana.bind_all("<Control-F4>", bloquear_alt_f4)

# --- INICIALIZACIÓN ---

# Verificar internet en segundo plano para no congelar el arranque
ventana.after(300, iniciar_conexion_en_segundo_plano)
ventana.after(1800, iniciar_actualizacion_en_segundo_plano)
ventana.after(5000, iniciar_verificacion_periodica)
ventana.after(
    600,
    lambda: iniciar_heartbeat_manager(
        DIRECTORIO_APP, VERSION_SISTEMA, registrar_evento_tecnico
    )
)
ventana.after_idle(notificar_interfaz_lista)
#ventana.after(1200, mostrar_instrucciones_iniciales, "")

ventana.mainloop()
