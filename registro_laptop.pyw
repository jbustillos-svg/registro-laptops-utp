import tkinter as tk
from tkinter import messagebox
import threading
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import socket
import requests
import pytz
import os
import time
import psutil

# Intentar importar PIL, si no está disponible usar emojis
try:
    from PIL import Image, ImageTk
    PIL_DISPONIBLE = True
except ImportError:
    PIL_DISPONIBLE = False

# --- VARIABLES GLOBALES ---
VERSION_SISTEMA = "v1.1.2-test"
hoja_alumnos = None
hoja_registros = None
zona_horaria = pytz.timezone("America/Chihuahua")
verificar_conexion_activo = True
procesando_sesion = False
aviso_internet = None
acepta_estado_equipo = None
chk_label = None

# =========================
# COLUMNAS HOJA REGISTROS (0-based)
# =========================
COL_MATRICULA = 0
COL_NOMBRE = 1
COL_FECHA = 2
COL_HORA_INGRESO = 3
COL_CONFIRMACION = 4
COL_HORA_SALIDA = 5
COL_LAPTOP_ID = 6
COL_BATERIA_ENTRADA = 7
COL_BATERIA_SALIDA = 8
# =========================
# 🎨 PALETA DE COLORES MODERNA
# =========================
COLOR_PRIMARIO = "#0066cc"          # Azul principal
COLOR_SECUNDARIO = "#00a8ff"       # Azul claro
COLOR_FONDO = "#f5f7fa"            # Gris muy claro
COLOR_TARJETA = "#ffffff"          # Blanco
COLOR_TEXTO = "#2d3436"            # Gris oscuro
COLOR_TEXTO_SECUNDARIO = "#636e72" # Gris medio
COLOR_EXITO = "#00b894"            # Verde
COLOR_ERROR = "#e84118"            # Rojo
COLOR_ADVERTENCIA = "#fdcb6e"      # Amarillo
COLOR_BORDE = "#dfe6e9"           # Borde gris claro
COLOR_HOVER = "#f1f2f6"           # Hover gris muy claro

# =========================
# FUENTES MODERNAS
# =========================
FUENTE_TITULO = ("Segoe UI", 24, "bold")
FUENTE_SUBTITULO = ("Segoe UI", 14, "bold")
FUENTE_CUERPO = ("Segoe UI", 11)
FUENTE_CUERPO_BOLD = ("Segoe UI", 11, "bold")
FUENTE_ENTRADA = ("Segoe UI", 14)
FUENTE_BOTON = ("Segoe UI", 12, "bold")
FUENTE_PEQ = ("Segoe UI", 9)
FUENTE_PEQ_BOLD = ("Segoe UI", 9, "bold")

def centrar_ventana(ventana, ancho, alto):
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


# --- FUNCIONES BASE (SIN CAMBIOS) ---
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
    ventana_ctrl.title("Aviso importante")
    ventana_ctrl.resizable(False, False)
    ventana_ctrl.configure(bg="#ffffff")
    
    # Aplicar estilo moderno
    ventana_ctrl.configure(bg=COLOR_FONDO)
    
    # Hacerla modal respecto a la ventana principal
    ventana_ctrl.transient(ventana)
    ventana_ctrl.grab_set()

    ancho = 500
    alto = 300 if estado == "SANCIONADO" else 420


    pantalla_ancho = ventana_ctrl.winfo_screenwidth()
    pantalla_alto = ventana_ctrl.winfo_screenheight()
    x = (pantalla_ancho - ancho) // 2
    y = (pantalla_alto - alto) // 2

    ventana_ctrl.geometry(f"{ancho}x{alto}+{x}+{y}")

    frame = tk.Frame(ventana_ctrl, bg=COLOR_TARJETA)
    frame.pack(fill=tk.BOTH, expand=True)

    contenido = tk.Frame(frame, bg=COLOR_TARJETA, padx=30, pady=12)
    contenido.pack(fill=tk.X)

    boton_frame = tk.Frame(frame, bg=COLOR_TARJETA, padx=40, pady=5)
    boton_frame.pack(side=tk.BOTTOM, fill=tk.X)


    # 🔴 USUARIO SANCIONADO
    if estado == "SANCIONADO":
        titulo = "🚫 USUARIO SANCIONADO"
        color = COLOR_ERROR
        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n\n"
            "Has excedido el número permitido.\n\n"
            "Acude con el administrador."
        )
    else:
        # ⚠️ AVISO GENERAL
        titulo = "⚠️ AVISO IMPORTANTE"
        color = COLOR_ADVERTENCIA

        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n"
            f"Estado: {estado}\n\n"
            "IMPORTANTE:\n"
            "Solo se permiten 4 intentos.\n"
            "Al superar este límite, el usuario será bloqueado."
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
        text="✔ ENTENDIDO",
        font=("Segoe UI", 13, "bold"),   # ⬅ más equilibrado
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
        padx=90,        # ⬅ más angosto
        pady=(8, 5),
        ipady=6         # ⬅ altura fina y controlada
    )





    # ⏸️ ESPERAR a que el usuario cierre la ventana
    ventana_ctrl.wait_window()

def incrementar_no_entregas(matricula):
    """
    Incrementa SOLO el contador de No_Entregas.
    El estado se calcula ÚNICAMENTE por fórmula en Google Sheets.
    """
    try:
        if hoja_alumnos is None:
            return 0, "ACTIVO"

        matriculas = hoja_alumnos.col_values(1)
        if matricula not in matriculas:
            return 0, "ACTIVO"

        fila = matriculas.index(matricula) + 1

        # Columna E = No_Entregas
        no_entregas_actual = hoja_alumnos.cell(fila, 5).value
        no_entregas_actual = int(no_entregas_actual) if no_entregas_actual else 0

        no_entregas_nuevo = no_entregas_actual + 1
        hoja_alumnos.update_cell(fila, 5, no_entregas_nuevo)

        # ⚠️ NO TOCAR columna del estado (tiene fórmula)
        return no_entregas_nuevo, "CALCULADO_POR_FORMULA"

    except Exception as e:
        print(f"Error al incrementar no entregas: {e}")
        return 0, "ACTIVO"
    
def cerrar_sesion_anterior_y_contar_no_entrega(matricula):
    """
    Cierra la sesión activa anterior del alumno,
    registra NO ENTREGA y deja evidencia.
    """
    try:
        if hoja_registros is None or hoja_alumnos is None:
            return False

        registros = hoja_registros.get_all_values()
        hora_actual, _ = obtener_hora_internet()

        for i in reversed(range(len(registros))):
            fila = registros[i]

            if fila[COL_MATRICULA] == matricula:
                if fila[COL_HORA_SALIDA].strip() == "":
                    hoja_registros.update_cell(i + 1, COL_HORA_SALIDA + 1, hora_actual)
                    hoja_registros.update_cell(
                        i + 1,
                        COL_BATERIA_SALIDA + 1,
                        "CIERRE_AUTOMATICO_POR_NUEVA_SESION"
                    )

                    incrementar_no_entregas(matricula)
                    return True
                break

    except Exception as e:
        print(f"Error al cerrar sesión anterior: {e}")

    return False


def obtener_control_alumno(matricula):
    """
    Obtiene el número de no entregas y el estado del alumno
    desde la hoja 'Alumnos'.
    """
    try:
        if hoja_alumnos is None:
            return 0, "ACTIVO"

        matriculas = hoja_alumnos.col_values(1)
        if matricula in matriculas:
            fila = matriculas.index(matricula) + 1

            no_entregas = hoja_alumnos.cell(fila, 5).value
            estado = hoja_alumnos.cell(fila, 7).value

            no_entregas = int(no_entregas) if no_entregas else 0
            estado = estado if estado else "ACTIVO"

            return no_entregas, estado

    except Exception as e:
        print(f"Error al obtener control del alumno: {e}")

    return 0, "ACTIVO"

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
    height=8,        # 👈 clave
    padx=20,
    pady=4
)
    texto_widget.pack(pady=6, fill=tk.X)

    texto_widget.insert(tk.END, texto)
    texto_widget.config(state=tk.DISABLED)




    tk.Button(
        frame,
        text="✔ ENTENDIDO",
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

def cerrar_sistema_admin(event=None):
    """
    Cierre administrativo del sistema mediante combinación de teclas.
    Uso exclusivo del administrador.
    """
    respuesta = messagebox.askyesno(
        "Cierre administrativo",
        "¿Desea cerrar el sistema?\n\nEsta acción es solo para el administrador."
    )

    if respuesta:
        try:
            detener_verificacion_conexion()
            ventana.destroy()
        except:
            os._exit(0)

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

def verificar_internet():
    try:
        requests.get("http://www.google.com", timeout=3)
        return True
    except:
        return False

def conectar_google_sheets():
    global hoja_alumnos, hoja_registros
    if not verificar_internet():
        cambiar_estado("🔴 Sin conexión a internet", COLOR_ERROR)
        return False
    
    cambiar_estado("🟡 Conectando...", COLOR_ADVERTENCIA)
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open("Control de Laptops")
        hoja_alumnos = sheet.worksheet("Alumnos")
        hoja_registros = sheet.worksheet("Registros")
        cambiar_estado("🟢 Conectado", COLOR_EXITO)
        return True
    except Exception as e:
        hoja_alumnos = None
        hoja_registros = None
        cambiar_estado("🔴 Error en la conexión", COLOR_ERROR)
        return False

def verificar_conexion_periodicamente():
    global verificar_conexion_activo, aviso_internet
    
    while verificar_conexion_activo:
        if verificar_internet():
            if hoja_alumnos is None or hoja_registros is None:
                if conectar_google_sheets():
                    # Cerrar aviso de internet si está abierto
                    if aviso_internet and aviso_internet.winfo_exists():
                        aviso_internet.destroy()
                        aviso_internet = None
                    
            else:
                cambiar_estado("🟢 Conectado", COLOR_EXITO)
        else:
            cambiar_estado("🔴 Sin conexión a internet", COLOR_ERROR)
            if hoja_alumnos is None or hoja_registros is None:
                conectar_google_sheets()
        
        time.sleep(2)

def detener_verificacion_conexion():
    global verificar_conexion_activo
    verificar_conexion_activo = False

def obtener_hora_internet():
    try:
        response = requests.get("http://worldtimeapi.org/api/timezone/America/Chihuahua", timeout=5)
        hora_utc = datetime.fromisoformat(response.json()["datetime"].split(".")[0])
        return zona_horaria.localize(hora_utc).strftime("%H:%M:%S"), zona_horaria.localize(hora_utc).strftime("%Y-%m-%d")
    except:
        ahora = datetime.now(zona_horaria)
        return ahora.strftime("%H:%M:%S"), ahora.strftime("%Y-%m-%d")

def obtener_porcentaje_bateria():
    try:
        bateria = psutil.sensors_battery()
        if bateria:
            return f"{int(bateria.percent)}%"
        else:
            return "N/A"
    except:
        return "N/A"

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
    ventana_curp = tk.Toplevel(parent)

    ventana_curp.transient(parent)
    ventana_curp.attributes("-topmost", True)
    ventana_curp.grab_set()
    ventana_curp.focus_force()
    ventana_curp.lift()

    ventana_curp.title("Validación de identidad")
    ventana_curp.resizable(False, False)

    ancho, alto = 360, 220
    x = (ventana_curp.winfo_screenwidth() - ancho) // 2
    y = (ventana_curp.winfo_screenheight() - alto) // 2
    ventana_curp.geometry(f"{ancho}x{alto}+{x}+{y}")

    tk.Label(
        ventana_curp,
        text="Por seguridad, ingresa los\nÚLTIMOS 2 DÍGITOS de tu CURP",
        font=("Segoe UI", 11),
        justify="center"
    ).pack(pady=20)

    resultado = {"valor": None}

    # 🔑 VARIABLE CONTROLADA
    curp_var = tk.StringVar()

    def controlar_texto(*args):
        valor = curp_var.get().upper()

        # Limitar a 2 caracteres
        if len(valor) > 2:
            valor = valor[:2]

        # Evitar loop infinito
        if curp_var.get() != valor:
            curp_var.set(valor)

    # 🔗 Vincular control
    curp_var.trace_add("write", controlar_texto)

    entrada = tk.Entry(
        ventana_curp,
        textvariable=curp_var,
        font=("Segoe UI", 16),
        justify="center",
        width=8,      # 👈 alargado, no gigante
    )
    entrada.pack(pady=10, ipady=6)
    entrada.focus()

    def confirmar(event=None):
        valor = curp_var.get().strip()

        if len(valor) != 2:
            messagebox.showerror(
                "Dato inválido",
                "Debes ingresar exactamente 2 caracteres.",
                parent=ventana_curp
            )
            return

        resultado["valor"] = valor
        ventana_curp.attributes("-topmost", False)
        ventana_curp.destroy()

    entrada.bind("<Return>", confirmar)

    tk.Button(
        ventana_curp,
        text="VALIDAR",
        font=FUENTE_BOTON,
        bg=COLOR_EXITO,
        fg="white",
        bd=0,
        padx=25,
        pady=10,
        command=confirmar
    ).pack(pady=20)

    ventana_curp.wait_window()
    return resultado["valor"]

def verificar_entrega_pendiente(matricula):
    try:
        if hoja_registros is None:
            return False, None, None

        registros = hoja_registros.get_all_values()

        for i in reversed(range(len(registros))):
            if registros[i][COL_MATRICULA] == matricula:
                hora_salida = registros[i][COL_HORA_SALIDA].strip()
                if hora_salida == "":
                    fecha_entrada = registros[i][COL_FECHA]
                    laptop_id = registros[i][COL_LAPTOP_ID]
                    return True, fecha_entrada, laptop_id
                break  # 👈 SOLO cuando ya encontró la matrícula

    except Exception as e:
        print(f"Error al verificar entrega pendiente: {e}")

    return False, None, None


def procesar_no_entrega_si_corresponde(matricula):
    try:
        if hoja_registros is None or hoja_alumnos is None:
            return

        registros = hoja_registros.get_all_values()

        for i in reversed(range(len(registros))):
            fila = registros[i]

            if fila[COL_MATRICULA] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                confirmacion = fila[COL_CONFIRMACION].strip()

                # 🔥 NO ENTREGA: confirmo entrada pero nunca entregó
                if hora_salida == "" and confirmacion == "CONFIRMADO":
                    incrementar_no_entregas(matricula)
                    hoja_registros.update_cell(
                        i + 1,
                        COL_BATERIA_SALIDA + 1,
                        "NO_ENTREGA_CONTADA"
                    )
                break

    except Exception as e:
        print(f"Error al procesar no entrega: {e}")


def sesion_activa_en_esta_laptop(matricula):
    try:
        if hoja_registros is None:
            return False

        laptop_actual = socket.gethostname()
        registros = hoja_registros.get_all_values()

        for fila in reversed(registros):
            if fila[COL_MATRICULA] == matricula:
                return (
                    fila[COL_LAPTOP_ID] == laptop_actual and
                    fila[COL_HORA_SALIDA].strip() == ""
                )

    except Exception as e:
        print(f"Error al validar sesión activa local: {e}")

    return False




def registrar_entrada(matricula):
    try:
        if hoja_registros is None:
            return None

        nombre = buscar_nombre(matricula)
        if nombre:
            hora, fecha = obtener_hora_internet()
            laptop_id = socket.gethostname()
            bateria_entrada = obtener_porcentaje_bateria()

            hoja_registros.append_row([
                matricula,          # A
                nombre,             # B
                fecha,              # C
                hora,               # D
                "CONFIRMADO",       # E  ← AQUÍ
                "",                 # F Hora_Salida
                laptop_id,          # G
                bateria_entrada,    # H
                ""                  # I Bateria_Salida
            ])

            return nombre
    except Exception as e:
        print(f"Error al registrar entrada: {e}")
    return None


def registrar_salida_con_reintentos(nombre, matricula, max_reintentos=5):
    laptop_actual = socket.gethostname()

    for intento in range(max_reintentos):
        try:
            if not verificar_internet():
                time.sleep(2)
                continue

            if hoja_registros is None:
                if not conectar_google_sheets():
                    time.sleep(2)
                    continue

            hora, _ = obtener_hora_internet()
            bateria_salida = obtener_porcentaje_bateria()
            registros = hoja_registros.get_all_values()

            for i in reversed(range(len(registros))):
                fila = registros[i]

                if (
                    fila[COL_MATRICULA] == matricula and
                    fila[COL_HORA_SALIDA] == "" and
                    fila[COL_LAPTOP_ID] == laptop_actual
                ):
                    hoja_registros.update_cell(i + 1, COL_HORA_SALIDA + 1, hora)
                    hoja_registros.update_cell(i + 1, COL_BATERIA_SALIDA + 1, bateria_salida)

                    return True

            return False

        except Exception as e:
            print(f"Error intento {intento + 1}: {e}")
            time.sleep(2)

    return False


def mostrar_ventana_espera_registro(ventana_entrega, matricula, nombre):
    """Muestra ventana de espera mientras se intenta registrar la salida"""
    
    ventana_espera = tk.Toplevel(ventana_entrega)
    ventana_espera.title("Registrando Salida")
    ventana_espera.resizable(False, False)
    ventana_espera.attributes('-topmost', True)
    ventana_espera.configure(bg=COLOR_FONDO)
    ANCHO = 440
    ALTO = 260
    centrar_ventana(ventana_espera, ANCHO, ALTO)
    ventana_espera.attributes('-topmost', True)

    
    # Centrar ventana
    ventana_espera.transient(ventana_entrega)
    ventana_espera.grab_set()
    
    frame = tk.Frame(ventana_espera, bg=COLOR_TARJETA, padx=25, pady=25)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Icono de carga
    tk.Label(frame, 
             text="🔄",
             font=("Segoe UI", 20),
             bg=COLOR_TARJETA).pack(pady=(0, 10))
    
    tk.Label(frame, 
             text="Registrando salida...",
             font=("Segoe UI", 12, "bold"),
             bg=COLOR_TARJETA).pack(pady=5)
    tk.Label(
    frame,
    text="Por favor espere mientras se guarda la información\nIntentando conexión a internet...",
    font=FUENTE_PEQ,
    fg=COLOR_TEXTO_SECUNDARIO,
    bg=COLOR_TARJETA,
    justify="center",
    wraplength=360
    ).pack(pady=5)

    
    # Barra de progreso
    progress_frame = tk.Frame(frame, bg=COLOR_TARJETA)
    progress_frame.pack(pady=15)
    
    progress_bar = tk.Frame(progress_frame, height=6, bg=COLOR_BORDE, width=300)
    progress_bar.pack()
    progress_bar_inner = tk.Frame(progress_bar, height=6, bg=COLOR_PRIMARIO, width=0)
    progress_bar_inner.place(relx=0, rely=0, relheight=1)
    
    estado_label = tk.Label(frame, 
                           text="Conectando...",
                           font=FUENTE_PEQ,
                           fg=COLOR_TEXTO_SECUNDARIO,
                           bg=COLOR_TARJETA)
    estado_label.pack(pady=5)
    
    def actualizar_progreso(progreso, mensaje):
        progress_bar_inner.config(width=progreso * 50)
        estado_label.config(text=mensaje)
        ventana_espera.update()
    
    def intentar_registro():
        for i in range(6):  # 0 a 5 (para 5 reintentos + éxito)
            if not ventana_espera.winfo_exists():
                return False
                
            if i < 5:
                actualizar_progreso(i, f"Intentando conexión ({i+1}/5)…")
                time.sleep(0.5)
            else:
                actualizar_progreso(5, "Registrando salida...")
                
            if registrar_salida_con_reintentos(nombre, matricula):
                actualizar_progreso(6, "✅ Salida registrada correctamente")
                time.sleep(1)
                ventana_espera.destroy()
                return True
            else:
                if i == 4:  # Último intento fallido
                    actualizar_progreso(5, "❌ Error al registrar salida")
                    time.sleep(2)
                    ventana_espera.destroy()
                    return False
        
        ventana_espera.destroy()
        return False
    
    # Ejecutar en un hilo separado
    def ejecutar_registro():
        exito = intentar_registro()
        if exito:
            ventana_entrega.destroy()
            os.system("shutdown /s /t 3")
        else:
            messagebox.showerror("Error", 
                               "No se pudo registrar la salida.\n\n" +
                               "La laptop no se apagará.\n" +
                               "Por favor intenta nuevamente.")
    
    threading.Thread(target=ejecutar_registro, daemon=True).start()
    
    return ventana_espera

def entregar_y_apagar(ventana, matricula, nombre):
    """
    Entrega correcta:
    - Si NO hay internet → NO apagar, NO sancionar
    - Si la sesión fue cerrada por otra laptop → aviso y apagado
    - Si todo está bien → registrar salida normal
    """

    # 🔴 1. PRIMERO validar conexión a internet
    if not verificar_internet():
        messagebox.showerror(
            "Sin conexión a internet",
            "No se pudo registrar la entrega porque no hay conexión a internet.\n\n"
            "Conéctate a internet y vuelve a intentar.\n\n"
            "La laptop NO se apagará."
        )
        return

    # 🔴 2. Validar sesión activa SOLO si hay conexión
    if not sesion_activa_en_esta_laptop(matricula):
        messagebox.showwarning(
            "Sesión cerrada",
            "Esta sesión ya fue cerrada automáticamente.\n\n"
            "El uso quedó registrado como NO ENTREGA.\n\n"
            "La computadora se apagará."
        )
        ventana.destroy()
        os.system("shutdown /s /t 3")
        return

    # 🟢 3. Flujo normal
    mostrar_ventana_espera_registro(ventana, matricula, nombre)




def mostrar_ventana_entrega(nombre, matricula):
    ventana_entrega = tk.Toplevel()
    
    # 🔑 CLAVE: ventana independiente
    ventana_entrega.transient(None)
    ventana_entrega.title("Entrega de Laptop")
    ventana_entrega.resizable(False, False)
    ventana_entrega.configure(bg=COLOR_FONDO)

    # ❌ Bloquear cerrar, pero NO minimizar
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

    # Logo
    logo_entrega = cargar_logo("UTP.png", 120, 70)
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
    tk.Button(
        frame_principal,
        text="🚀 ENTREGAR Y APAGAR",
        font=FUENTE_BOTON,
        bg=COLOR_ERROR,
        fg="white",
        padx=30,
        pady=15,
        bd=0,
        cursor="hand2",
        activebackground="#c23616",
        command=lambda: entregar_y_apagar(ventana_entrega, matricula, nombre)
    ).pack(pady=20)

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
    
    mensaje = f"⚠️ Olvidaste entregar la laptop\n\n"
    mensaje += f"Laptop: {laptop_id}\n"
    mensaje += f"Fecha: {fecha_formateada}\n\n"
    mensaje += "Recuerda siempre usar el botón 'Entregar y Apagar'"
    
    messagebox.showwarning("Entrega Pendiente", mensaje)
def verificar_sesion_activa_en_otra_laptop(matricula):
    """
    Verifica si la matrícula tiene una sesión activa
    en OTRA laptop diferente a la actual.
    """
    try:
        if hoja_registros is None:
            return False, None

        laptop_actual = socket.gethostname()
        registros = hoja_registros.get_all_values()

        for fila in reversed(registros):
            if fila[0] == matricula:
                hora_salida = fila[COL_HORA_SALIDA].strip()
                laptop_registro = fila[COL_LAPTOP_ID]


                if hora_salida == "" and laptop_registro != laptop_actual:
                    return True, laptop_registro
                break

    except Exception as e:
        print(f"Error al verificar sesión en otra laptop: {e}")

    return False, None


def mostrar_confirmacion_simple(nombre, matricula):
    global entrada, ventana, procesando_sesion

    # 🔓 LIBERAR ESTADO SIEMPRE AL ENTRAR AQUÍ
    procesando_sesion = False
    btn_entrar.config(state="normal")


    # 1. Confirmar identidad
    respuesta = messagebox.askyesno(
        "Confirmación",
        f"¿Eres {nombre}?"
    )

    if not respuesta:
        reiniciar_estado_sistema()
        entrada.delete(0, tk.END)
        entrada.focus()
        return False

    
    # 🔒 VERIFICAR SESIÓN ACTIVA EN OTRA LAPTOP
    sesion_activa, laptop_otro = verificar_sesion_activa_en_otra_laptop(matricula)

    if sesion_activa:
        respuesta = messagebox.askyesno(
            "Sesión activa detectada",
            "Se detectó una sesión activa en otro equipo.\n\n"
            f"Laptop anterior: {laptop_otro}\n\n"
            "Si continúas:\n"
            "• La sesión anterior se cerrará automáticamente\n"
            "• Se registrará como NO ENTREGA\n"
            "• Esta acción quedará registrada\n\n"
            "¿Deseas continuar en este equipo?"
        )

        if not respuesta:
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return False


        # 🔥 Cerrar sesión anterior y contar NO ENTREGA
        exito = cerrar_sesion_anterior_y_contar_no_entrega(matricula)

        if not exito:
            messagebox.showerror(
                "Error",
                "No se pudo cerrar la sesión anterior.\n\nIntenta nuevamente."
            )
            reiniciar_estado_sistema()
            return False

    # 🔐 VALIDAR ROL
    rol = buscar_rol(matricula)

    # 🔒 VALIDACIÓN POR CURP SOLO PARA ALUMNOS
    if rol == "ALUMNO":

        curp_real = buscar_curp(matricula)

        if not curp_real:
            messagebox.showerror(
                "Error",
                "No se pudo validar la identidad.\nContacta al administrador."
            )
            reiniciar_estado_sistema()
            return False

        curp_ingresada = pedir_curp_ultimos_2(ventana)

        if not curp_ingresada:
            reiniciar_estado_sistema()
            return False

        if not validar_curp_ultimos_2(curp_real, curp_ingresada):
            messagebox.showerror(
                "Acceso denegado",
                "Los datos no coinciden.\n\nSesión cancelada."
            )
            reiniciar_estado_sistema()
            return False


    # 2. 🔥 CONTAR NO ENTREGA AUTOMÁTICA (AQUÍ ES DONDE IBA ANTES)
    procesar_no_entrega_si_corresponde(matricula)

    # 3. Obtener control actualizado
    no_entregas, estado = obtener_control_alumno(matricula)

    # 4. Verificar entrega pendiente actual
    tiene_pendiente, fecha_entrada, laptop_id = verificar_entrega_pendiente(matricula)

    # 5. Mostrar aviso SOLO si tiene 2 o más
    if no_entregas >= 2:
        mostrar_ventana_control_unificada(
            matricula=matricula,
            nombre=nombre,
            tiene_pendiente=tiene_pendiente,
            fecha_entrada=fecha_entrada,
            laptop_id=laptop_id,
            no_entregas=no_entregas,
            estado=estado
        )

        if estado == "SANCIONADO":
            reiniciar_estado_sistema()
            entrada.delete(0, tk.END)
            entrada.focus()
            return False


    # 6. Registrar nueva entrada
    resultado = registrar_entrada(matricula)

    if not resultado:
        messagebox.showerror(
            "Error",
            "No se pudo registrar la entrada."
        )
        reiniciar_estado_sistema()
        return False


    # 7. Ocultar ventana principal
    ventana.withdraw()

    # 8. Abrir ventana de entrega
    ventana.after(
        200,
        lambda: mostrar_ventana_entrega(nombre, matricula)
    )
    return True


def cambiar_estado(texto, color=None):
    estado_var.set(texto)
    if color:
        estado_label.config(fg=color)
    ventana.update_idletasks()

def verificar_conexion_base_datos():
    """Verifica si hay conexión a la base de datos (Google Sheets)"""
    if hoja_alumnos is None or hoja_registros is None:
        return False
    try:
        # Intentar una operación simple para verificar conexión
        hoja_alumnos.row_count
        return True
    except:
        return False

def reiniciar_estado_sistema():
    """Reinicia el estado del sistema después de una operación"""
    global procesando_sesion
    procesando_sesion = False
    btn_entrar.config(state="normal")
    # Actualizar estado de conexión actual
    if verificar_conexion_base_datos():
        cambiar_estado("🟢 Conectado", COLOR_EXITO)
    else:
        cambiar_estado("🔴 Sin conexión", COLOR_ERROR)
        if acepta_estado_equipo:
            acepta_estado_equipo.set(False)
            try:
                chk_label.config(text="⬜ ")
            except:
                pass



def mostrar_aviso_internet_bloqueante():
    """Muestra un aviso emergente bloqueante que no se puede mover, cerrar ni minimizar"""
    global aviso_internet
    
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
    logo_aviso = cargar_logo("UTP.png", 70, 45)
    if logo_aviso:
        lbl_logo_aviso = tk.Label(frame_aviso, image=logo_aviso, bg=COLOR_TARJETA)
        lbl_logo_aviso.image = logo_aviso
        lbl_logo_aviso.pack(pady=(25, 10))
    
    # Icono de advertencia
    tk.Label(frame_aviso, 
             text="🔴",
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
        
        progress_bar_inner.config(width=progreso_actual)
        aviso_internet.after(50, animar_barra_progreso)
    
    # Función para verificar internet continuamente
    def verificar_internet_continuamente():
        if verificar_internet():
            # Internet detectado, cerrar aviso
            aviso_internet.destroy()
            cambiar_estado("🟢 Conectado", COLOR_EXITO)
            # Intentar conectar a Google Sheets
            threading.Thread(target=conectar_google_sheets, daemon=True).start()
        else:
            # Actualizar tiempo de espera
            tiempo_transcurrido = int(time.time() - tiempo_inicio)
            tiempo_label.config(text=f"Tiempo esperando: {tiempo_transcurrido} segundos")
            
            # Seguir verificando cada segundo
            estado_label.config(text=f"Verificando conexión... ({time.strftime('%H:%M:%S')})")
            aviso_internet.after(1000, verificar_internet_continuamente)
    
    # Iniciar animación y verificación
    progreso_actual = 0
    direccion = 1
    animar_barra_progreso()
    aviso_internet.after(500, verificar_internet_continuamente)
def iniciar_sesion():
    global procesando_sesion

    if procesando_sesion:
        return


    procesando_sesion = True
    btn_entrar.config(state="disabled")
    cambiar_estado("🟡 Verificando conexión...", COLOR_ADVERTENCIA)
    ventana.update()

    matricula = entrada.get().strip()

    if not matricula:
        messagebox.showwarning("Campo vacío", "Por favor ingresa tu matrícula.")
        entrada.focus()
        reiniciar_estado_sistema()
        return
        # 🔒 Validar aceptación de estado del equipo
    if not acepta_estado_equipo.get():
        messagebox.showwarning(
            "Confirmación requerida",
            "Debes confirmar el estado de la laptop para continuar."
        )
        reiniciar_estado_sistema()
        return


    if not verificar_internet():
        messagebox.showerror(
            "Sin conexión a internet",
            "No hay conexión a internet.\n\nConecte la computadora e intente nuevamente."
        )
        reiniciar_estado_sistema()
        return

    if not verificar_conexion_base_datos():
        messagebox.showerror(
            "Error de conexión",
            "No hay conexión a la base de datos."
        )
        reiniciar_estado_sistema()
        return

    nombre = buscar_nombre(matricula)

    if not nombre:
        messagebox.showerror(
            "Matrícula no válida",
            "La matrícula no está registrada.\n\nVerifica o contacta al administrador."
        )
        entrada.delete(0, tk.END)
        entrada.focus()
        reiniciar_estado_sistema()
        return

    resultado = mostrar_confirmacion_simple(nombre, matricula)

    if resultado is False:
        reiniciar_estado_sistema()




def crear_pantalla_login():
    global entrada, btn_entrar, estado_var, estado_label, chk_label

    for widget in ventana.winfo_children():
        widget.destroy()

    ventana.configure(bg=COLOR_FONDO)

    def forzar_mayusculas(event):
        texto = entrada.get()
        entrada.delete(0, tk.END)
        entrada.insert(0, texto.upper())

    main_container = tk.Frame(ventana, bg=COLOR_FONDO)
    main_container.pack(fill=tk.BOTH, expand=True)

    card = tk.Frame(
        main_container,
        bg=COLOR_TARJETA,
        width=460,
        height=620,
        highlightthickness=1,
        highlightbackground=COLOR_BORDE
    )
    card.place(relx=0.5, rely=0.5, anchor="center")
    card.pack_propagate(False)

    # LOGO
    logo = cargar_logo("UTP.png", 150, 95)
    if logo:
        lbl_logo = tk.Label(card, image=logo, bg=COLOR_TARJETA)
        lbl_logo.image = logo
        lbl_logo.pack(pady=(35, 15))

    tk.Label(
        card,
        text="SISTEMA DE CONTROL\nDE LAPTOPS",
        font=("Segoe UI", 18, "bold"),
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
        font=("Segoe UI", 16),
        justify="center",
        bg="#f8f9fa",
        fg=COLOR_TEXTO,
        highlightthickness=2,
        highlightbackground=COLOR_BORDE,
        highlightcolor=COLOR_PRIMARIO
    )
    entrada.pack(fill=tk.X, ipady=12, pady=(0, 18))
    entrada.focus()

        # =========================
        # CONFIRMACIÓN ESTADO LAPTOP
        # =========================
    global acepta_estado_equipo
    acepta_estado_equipo = tk.BooleanVar(value=False)

    check_frame = tk.Frame(card, bg=COLOR_TARJETA)
    check_frame.pack(padx=50, pady=(0, 25), fill=tk.X)

        # Checkbutton con emoji que cambia
    def update_check_emoji():
            if acepta_estado_equipo.get():
                chk_label.config(text="✅ ")
            else:
                chk_label.config(text="⬜ ")
        
        # Frame para alinear checkbox y texto
    check_content = tk.Frame(check_frame, bg=COLOR_TARJETA)
    check_content.pack(anchor="w")

        # Label con emoji (hace las veces de checkbox)
    chk_label = tk.Label(
            check_content,
            text="⬜ ",
            font=("Segoe UI", 14),
            fg=COLOR_PRIMARIO,
            bg=COLOR_TARJETA,
            cursor="hand2"
        )
    chk_label.pack(side=tk.LEFT, padx=(0, 10))
    update_check_emoji()

        # Texto del checkbox
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

        # Hacer ambos elementos clickeables
    def toggle_check(event=None):
            acepta_estado_equipo.set(not acepta_estado_equipo.get())
            update_check_emoji()

    chk_label.bind("<Button-1>", toggle_check)
    chk_text.bind("<Button-1>", toggle_check)

        # Línea divisoria
    tk.Frame(check_frame, height=1, bg=COLOR_BORDE).pack(fill=tk.X, pady=(15, 0))

  
    entrada.bind("<KeyRelease>", forzar_mayusculas)
    entrada.bind("<Return>", lambda e: iniciar_sesion())


    # =========================
    # BOTÓN INGRESAR
    # =========================
    btn_entrar = tk.Button(
        card,
        text="INGRESAR AL SISTEMA",
        font=FUENTE_BOTON,
        bg=COLOR_PRIMARIO,
        fg="white",
        activebackground=COLOR_SECUNDARIO,
        bd=0,
        padx=20,
        pady=14,
        cursor="hand2",
        command=iniciar_sesion
    )
    btn_entrar.pack(fill=tk.X, padx=50, pady=(5, 20))

    # =========================
    # ESTADO DE CONEXIÓN
    # =========================
    estado_var = tk.StringVar(value="🟢 Conectado")
    estado_label = tk.Label(
        card,
        textvariable=estado_var,
        font=("Segoe UI", 10, "bold"),
        fg=COLOR_EXITO,
        bg=COLOR_TARJETA
    )
    estado_label.pack(pady=(5, 10))

    # =========================
    # VERSIÓN (ABAJO DEL TODO)
    # =========================
    tk.Label(
        card,
        text=f"Versión {VERSION_SISTEMA}",
        font=("Segoe UI", 9),
        fg=COLOR_TEXTO_SECUNDARIO,
        bg=COLOR_TARJETA
    ).pack(side=tk.BOTTOM, anchor="e", padx=12, pady=(5, 8))

# --- VENTANA PRINCIPAL ---
ventana = tk.Tk()
ventana.title(f"SISTEMA DE CONTROL DE LAPTOPS - UTP | {VERSION_SISTEMA}")

# Pantalla completa sin bordes
ventana.attributes("-topmost", True)
ventana.state("zoomed")
ventana.overrideredirect(True)
ventana.configure(bg=COLOR_FONDO)

# 👉 AQUÍ SE CREA EL LOGIN
crear_pantalla_login()

# SEGURIDAD
ventana.protocol("WM_DELETE_WINDOW", lambda: detener_verificacion_conexion() or ventana.destroy())
ventana.bind("<Control-Alt-u>", cerrar_sistema_admin)
ventana.bind("<Control-Alt-U>", cerrar_sistema_admin)
ventana.bind_all("<Alt-F4>", bloquear_alt_f4)
ventana.bind_all("<Control-F4>", bloquear_alt_f4)

# --- INICIALIZACIÓN ---

# Verificar internet al iniciar
if not verificar_internet():
    # Mostrar aviso bloqueante si no hay internet
    ventana.after(1000, mostrar_aviso_internet_bloqueante)
else:
    # Si hay internet, conectar normalmente
    threading.Thread(target=conectar_google_sheets, daemon=True).start()

threading.Thread(target=verificar_conexion_periodicamente, daemon=True).start()
ventana.after(1200, mostrar_instrucciones_iniciales, "")

ventana.mainloop()