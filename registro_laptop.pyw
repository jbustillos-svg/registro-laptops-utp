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
VERSION_SISTEMA = "v1.0.2"
hoja_alumnos = None
hoja_registros = None
zona_horaria = pytz.timezone("America/Chihuahua")
verificar_conexion_activo = True
mensaje_espera = None
procesando_sesion = False
aviso_internet = None

# =========================
# 🎨 ESTILOS GENERALES (UI)
# =========================

COLOR_FONDO_APP = "#0f3554"       # Azul institucional
COLOR_TARJETA = "#ffffff"
COLOR_TEXTO = "#2c3e50"
COLOR_SECUNDARIO = "#7f8c8d"
COLOR_PRIMARIO = "#1abc9c"
COLOR_PELIGRO = "#e74c3c"
COLOR_LINEA = "#3498db"

FUENTE_TITULO = ("Segoe UI", 18, "bold")
FUENTE_SUBTITULO = ("Segoe UI", 13, "bold")
FUENTE_TEXTO = ("Segoe UI", 11)
FUENTE_BOTON = ("Segoe UI", 12, "bold")
FUENTE_PIE = ("Segoe UI", 8)


# --- FUNCIONES BASE ---
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

    # Hacerla modal respecto a la ventana principal
    ventana_ctrl.transient(ventana)
    ventana_ctrl.grab_set()

    ancho = 480
    alto = 420

    pantalla_ancho = ventana_ctrl.winfo_screenwidth()
    pantalla_alto = ventana_ctrl.winfo_screenheight()
    x = (pantalla_ancho - ancho) // 2
    y = (pantalla_alto - alto) // 2

    ventana_ctrl.geometry(f"{ancho}x{alto}+{x}+{y}")

    frame = tk.Frame(ventana_ctrl, bg="#ffffff", padx=25, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # 🔴 USUARIO SANCIONADO
    if estado == "SANCIONADO":
        titulo = "🚫 USUARIO SANCIONADO"
        color = "#c0392b"
        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n\n"
            "Has excedido el número permitido.\n\n"
            "Acude con el administrador."
        )
    else:
        # ⚠️ AVISO GENERAL
        titulo = "⚠️ AVISO IMPORTANTE"
        color = "#f39c12"

        mensaje = (
            f"No entregas registradas: {no_entregas} de 4\n"
            f"Estado: {estado}\n\n"
            "IMPORTANTE:\n"
            "Solo se permiten 4 intentos.\n"
            "Al superar este límite, el usuario será bloqueado.\n\n"
        )

        if tiene_pendiente:
            mensaje = (
                "Entrega pendiente detectada\n\n"
                f"Laptop: {laptop_id}\n"
                f"Fecha: {formatear_fecha(fecha_entrada)}\n\n"
            ) + mensaje

    tk.Label(
        frame,
        text=titulo,
        font=("Arial", 14, "bold"),
        fg=color,
        bg="#ffffff"
    ).pack(pady=(0, 15))

    tk.Label(
        frame,
        text=mensaje,
        font=("Arial", 11),
        fg="#2c3e50",
        bg="#ffffff",
        justify=tk.CENTER,
        wraplength=420
    ).pack(pady=10)

    tk.Button(
        frame,
        text="ENTENDIDO",
        font=("Arial", 11, "bold"),
        bg="#3498db",
        fg="white",
        padx=25,
        pady=8,
        command=ventana_ctrl.destroy
    ).pack(pady=20)

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

        # Columna C = No_Entregas
        no_entregas_actual = hoja_alumnos.cell(fila, 3).value
        no_entregas_actual = int(no_entregas_actual) if no_entregas_actual else 0

        no_entregas_nuevo = no_entregas_actual + 1
        hoja_alumnos.update_cell(fila, 3, no_entregas_nuevo)

        # ⚠️ NO TOCAR columna del estado (tiene fórmula)
        return no_entregas_nuevo, "CALCULADO_POR_FORMULA"

    except Exception as e:
        print(f"Error al incrementar no entregas: {e}")
        return 0, "ACTIVO"


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

            no_entregas = hoja_alumnos.cell(fila, 3).value
            estado = hoja_alumnos.cell(fila, 5).value

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
    ventana_info.configure(bg="#ffffff")
    ventana_info.attributes('-topmost', True)

    ancho = 500
    alto = 360

    x = (ventana_info.winfo_screenwidth() - ancho) // 2
    y = (ventana_info.winfo_screenheight() - alto) // 2
    ventana_info.geometry(f"{ancho}x{alto}+{x}+{y}")

    ventana_info.transient(ventana)
    ventana_info.grab_set()

    frame = tk.Frame(ventana_info, bg="#ffffff", padx=25, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        frame,
        text="INSTRUCCIONES IMPORTANTES",
        font=("Arial", 14, "bold"),
        fg="#2c3e50",
        bg="#ffffff"
    ).pack(pady=(0, 15))

    texto = (
        "• Ingresa tu matrícula correctamente.\n\n"
        "• NO cierres el sistema ni apagues la computadora manualmente.\n\n"
        "• Al finalizar, utiliza siempre el botón:\n"
        "  “ENTREGAR Y APAGAR”.\n\n"
        "• El mal uso del equipo es responsabilidad del usuario."
    )

    tk.Label(
        frame,
        text=texto,
        font=("Arial", 11),
        fg="#34495e",
        bg="#ffffff",
        justify=tk.LEFT,
        wraplength=440
    ).pack(pady=10)

    tk.Button(
        frame,
        text="✔ ENTENDIDO",
        font=("Arial", 11, "bold"),
        bg="#27ae60",
        fg="white",
        padx=25,
        pady=8,
        command=ventana_info.destroy
    ).pack(pady=15)



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
        cambiar_estado("🔴 Sin conexión a internet")
        return False
    
    cambiar_estado("🟡 Conectando...")
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open("Control de Laptops")
        hoja_alumnos = sheet.worksheet("Alumnos")
        hoja_registros = sheet.worksheet("Registros")
        cambiar_estado("🟢 Conectado")
        return True
    except Exception as e:
        hoja_alumnos = None
        hoja_registros = None
        cambiar_estado("🔴 Error en la conexión")
        return False

def verificar_conexion_periodicamente():
    global verificar_conexion_activo, mensaje_espera, aviso_internet
    
    while verificar_conexion_activo:
        if verificar_internet():
            if hoja_alumnos is None or hoja_registros is None:
                if conectar_google_sheets():
                    # Cerrar aviso de internet si está abierto
                    if aviso_internet and aviso_internet.winfo_exists():
                        aviso_internet.destroy()
                        aviso_internet = None
                    if mensaje_espera:
                        mensaje_espera.destroy()
                        mensaje_espera = None
            else:
                cambiar_estado("🟢 Conectado")
        else:
            cambiar_estado("🔴 Sin conexión a internet")
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

def verificar_entrega_pendiente(matricula):
    try:
        if hoja_registros is None:
            return False, None, None
            
        registros = hoja_registros.get_all_values()
        for i in reversed(range(len(registros))):
            if registros[i][0] == matricula:
                hora_salida = registros[i][4].strip()
                if hora_salida == "":
                    fecha_entrada = registros[i][2]
                    laptop_id = registros[i][5] if len(registros[i]) > 5 else "N/A"
                    return True, fecha_entrada, laptop_id
                break
    except Exception as e:
        print(f"Error al verificar entrega pendiente: {e}")
    return False, None, None

def procesar_no_entrega_si_corresponde(matricula):
    """
    Incrementa No_Entregas SOLO UNA VEZ si hay
    una entrega pendiente no procesada.
    """
    try:
        if hoja_registros is None or hoja_alumnos is None:
            return

        registros = hoja_registros.get_all_values()

        for i in reversed(range(len(registros))):
            fila = registros[i]

            if fila[0] == matricula:
                hora_salida = fila[4].strip()
                observacion = fila[7].strip() if len(fila) > 7 else ""

                # Entrega pendiente NO contada
                if hora_salida == "" and observacion != "NO_ENTREGA_CONTADA":
                    incrementar_no_entregas(matricula)
                    hoja_registros.update_cell(i + 1, 8, "NO_ENTREGA_CONTADA")
                break

    except Exception as e:
        print(f"Error al procesar no entrega: {e}")

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
                matricula, 
                nombre, 
                fecha, 
                hora, 
                "",
                laptop_id, 
                bateria_entrada, 
                ""
            ])
            return nombre
    except Exception as e:
        print(f"Error al registrar entrada: {e}")
    return None

def registrar_salida_con_reintentos(nombre, matricula, max_reintentos=5):
    """Registra la salida con sistema de reintentos en caso de desconexión"""
    
    for intento in range(max_reintentos):
        try:
            # Verificar conexión antes de intentar
            if not verificar_internet():
                cambiar_estado("🔴 Sin conexión - Reintentando...")
                time.sleep(2)  # Esperar antes de reintentar
                continue
                
            # Reconectar si es necesario
            if hoja_registros is None:
                if not conectar_google_sheets():
                    time.sleep(2)
                    continue
            
            # Registrar salida
            hora, _ = obtener_hora_internet()
            bateria_salida = obtener_porcentaje_bateria()
            
            registros = hoja_registros.get_all_values()
            for i in reversed(range(len(registros))):
                if registros[i][0] == matricula and registros[i][4] == "":
                    hoja_registros.update_cell(i+1, 5, hora)
                    hoja_registros.update_cell(i+1, 8, bateria_salida)
                    cambiar_estado("🟢 Salida registrada correctamente")
                    return True  # Éxito
            
            # Si llegamos aquí, no se encontró registro pendiente
            print("No se encontró registro pendiente para la matrícula")
            return True
            
        except Exception as e:
            print(f"Error en intento {intento + 1}: {e}")
            if intento < max_reintentos - 1:  # No esperar en el último intento
                time.sleep(2)
    
    # Si fallan todos los reintentos
    cambiar_estado("🔴 Error al registrar salida")
    return False

def mostrar_ventana_espera_registro(ventana_entrega, matricula, nombre):
    """Muestra ventana de espera mientras se intenta registrar la salida"""
    
    ventana_espera = tk.Toplevel(ventana_entrega)
    ventana_espera.title("Registrando Salida")
    ventana_espera.resizable(False, False)
    ventana_espera.geometry("350x200")
    ventana_espera.configure(bg="#ffffff")
    ventana_espera.attributes('-topmost', True)
    
    # Centrar ventana
    ventana_espera.transient(ventana_entrega)
    ventana_espera.grab_set()
    
    frame = tk.Frame(ventana_espera, bg="#ffffff", padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    tk.Label(frame, 
             text="🔄 Registrando salida...",
             font=("Arial", 12, "bold"),
             bg="#ffffff").pack(pady=10)
    
    tk.Label(frame, 
             text="Por favor espere mientras se guarda la información",
             font=("Arial", 9),
             bg="#ffffff",
             wraplength=300).pack(pady=5)
    
    # Barra de progreso simulada
    progress_frame = tk.Frame(frame, bg="#ffffff")
    progress_frame.pack(pady=15)
    
    progress_bar = tk.Frame(progress_frame, height=8, bg="#ecf0f1", width=250)
    progress_bar.pack()
    progress_bar_inner = tk.Frame(progress_bar, height=8, bg="#3498db", width=0)
    progress_bar_inner.place(relx=0, rely=0, relheight=1)
    
    estado_label = tk.Label(frame, 
                           text="Conectando...",
                           font=("Arial", 9),
                           bg="#ffffff")
    estado_label.pack(pady=5)
    
    btn_cancelar = tk.Button(frame,
                            text="Cancelar",
                            font=("Arial", 9),
                            command=lambda: ventana_espera.destroy())
    btn_cancelar.pack(pady=10)
    
    def actualizar_progreso(progreso, mensaje):
        progress_bar_inner.config(width=progreso * 2.5)
        estado_label.config(text=mensaje)
        ventana_espera.update()
    
    def intentar_registro():
        for i in range(6):  # 0 a 5 (para 5 reintentos + éxito)
            if not ventana_espera.winfo_exists():
                return False
                
            if i < 5:
                actualizar_progreso(i, f"Intentando conexión... ({i+1}/5)")
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
    """Función mejorada con manejo de errores de conexión"""
    
    # Mostrar ventana de espera mientras se registra
    mostrar_ventana_espera_registro(ventana, matricula, nombre)

# --- INTERFAZ GRAFICA ---

class RoundedEntry(tk.Entry):
    def __init__(self, master=None, **kwargs):
        tk.Entry.__init__(self, master, **kwargs)
        self.config(relief='flat', bg='#ecf0f1')

class RoundedButton(tk.Button):
    def __init__(self, master=None, **kwargs):
        tk.Button.__init__(self, master, **kwargs)
        self.config(relief='flat', bd=0, cursor='hand2')

def mostrar_ventana_entrega(nombre, matricula):
    ventana_entrega = tk.Toplevel()
    
    # 🔑 CLAVE: ventana independiente
    ventana_entrega.transient(None)
    ventana_entrega.title("Entrega de Laptop")
    ventana_entrega.resizable(False, False)

    # ❌ Bloquear cerrar, pero NO minimizar
    ventana_entrega.protocol("WM_DELETE_WINDOW", lambda: None)

    ventana_entrega.configure(bg="#ffffff")

    # ❌ NO topmost permanente
    # (esto es lo que rompía minimizar)
    ventana_entrega.attributes('-topmost', False)

    # Permitir minimizar
    ventana_entrega.minsize(380, 400)

    frame_principal = tk.Frame(
        ventana_entrega,
        bg="#ffffff",
        padx=30,
        pady=25
    )
    frame_principal.pack(fill=tk.BOTH, expand=True)

    # Logo
    logo_entrega = cargar_logo("UTP.png", 100, 60)
    if logo_entrega:
        lbl_logo = tk.Label(frame_principal, image=logo_entrega, bg="#ffffff")
        lbl_logo.image = logo_entrega
        lbl_logo.pack(pady=(0, 15))

    tk.Label(
        frame_principal,
        text="Sistema de Control de Laptops",
        font=("Arial", 12, "bold"),
        bg="#ffffff"
    ).pack(pady=(0, 15))

    tk.Label(
        frame_principal,
        text=f"Bienvenid@, {nombre}",
        font=("Arial", 14, "bold"),
        fg="#27ae60",
        bg="#ffffff"
    ).pack(pady=(0, 20))

    RoundedButton(
        frame_principal,
        text="🚀 Entregar y Apagar",
        font=("Arial", 12, "bold"),
        bg="#e74c3c",
        fg="white",
        padx=25,
        pady=10,
        command=lambda: entregar_y_apagar(ventana_entrega, matricula, nombre)
    ).pack(pady=15)

    tk.Label(
        frame_principal,
        text="La laptop se apagará automáticamente",
        font=("Arial", 9),
        fg="#7f8c8d",
        bg="#ffffff"
    ).pack(pady=(10, 0))

    # Centrar ventana
    ancho, alto = 380, 400
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

def mostrar_confirmacion_simple(nombre, matricula):
    global entrada
    global ventana

    # 1. Confirmar identidad
    respuesta = messagebox.askyesno(
        "Confirmación",
        f"¿Eres {nombre}?"
    )

    if not respuesta:
        reiniciar_estado_sistema()
        entrada.delete(0, tk.END)
        entrada.focus()
        return

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
            return

    # 6. Registrar nueva entrada
    resultado = registrar_entrada(matricula)

    if not resultado:
        messagebox.showerror(
            "Error",
            "No se pudo registrar la entrada."
        )
        reiniciar_estado_sistema()
        return

    # 7. Ocultar ventana principal
    ventana.withdraw()

    # 8. Abrir ventana de entrega
    ventana.after(
        200,
        lambda: mostrar_ventana_entrega(nombre, matricula)
    )


def cambiar_estado(texto):
    estado_var.set(texto)
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
        cambiar_estado("🟢 Conectado")
    else:
        cambiar_estado("🔴 Sin conexión")

def mostrar_aviso_internet_bloqueante():
    """Muestra un aviso emergente bloqueante que no se puede mover, cerrar ni minimizar"""
    global aviso_internet
    
    aviso_internet = tk.Toplevel(ventana)
    aviso_internet.title("Conexión Requerida")
    
    # Hacer la ventana completamente bloqueante
    aviso_internet.attributes('-topmost', True)  # Siempre encima
    aviso_internet.resizable(False, False)  # No redimensionable
    aviso_internet.overrideredirect(True)  # Sin barra de título (no se puede mover, cerrar, minimizar)
    
    # Configurar tamaño y posición (centrada)
    ancho_aviso = 500
    alto_aviso = 300
    
    pantalla_ancho = aviso_internet.winfo_screenwidth()
    pantalla_alto = aviso_internet.winfo_screenheight()
    x = (pantalla_ancho - ancho_aviso) // 2
    y = (pantalla_alto - alto_aviso) // 2
    
    aviso_internet.geometry(f"{ancho_aviso}x{alto_aviso}+{x}+{y}")
    aviso_internet.configure(bg="#1a5276")
    
    # Frame principal
    frame_aviso = tk.Frame(aviso_internet, bg="#ffffff", bd=2, relief=tk.RAISED)
    frame_aviso.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # Logo
    logo_aviso = cargar_logo("UTP.png", 80, 50)
    if logo_aviso:
        lbl_logo_aviso = tk.Label(frame_aviso, image=logo_aviso, bg="#ffffff")
        lbl_logo_aviso.image = logo_aviso
        lbl_logo_aviso.pack(pady=(20, 10))
    else:
        tk.Label(frame_aviso, text="🏢", font=("Arial", 24), bg="#ffffff").pack(pady=(20, 5))
        tk.Label(frame_aviso, text="UTP", font=("Arial", 12, "bold"), fg="#2c3e50", bg="#ffffff").pack()
    
    # Icono de advertencia
    tk.Label(frame_aviso, 
             text="🔴",
             font=("Arial", 28),
             bg="#ffffff").pack(pady=(10, 15))
    
    # Mensaje principal
    tk.Label(frame_aviso, 
             text="NO SE HA DETECTADO CONEXIÓN A INTERNET",
             font=("Arial", 12, "bold"),
             fg="#e74c3c",
             bg="#ffffff").pack(pady=(0, 10))
    
    # Instrucciones
    tk.Label(frame_aviso, 
             text="Para utilizar el sistema es necesario contar con conexión a internet.",
             font=("Arial", 9),
             bg="#ffffff",
             wraplength=400,
             justify=tk.CENTER).pack(pady=(0, 5))
    
    tk.Label(frame_aviso, 
             text="Conecte la computadora a internet para continuar.",
             font=("Arial", 9, "bold"),
             bg="#ffffff",
             wraplength=400,
             justify=tk.CENTER).pack(pady=(0, 15))
    
    # Barra de progreso animada
    progress_frame = tk.Frame(frame_aviso, bg="#ffffff")
    progress_frame.pack(pady=10)
    
    progress_bar = tk.Frame(progress_frame, height=6, bg="#ecf0f1", width=350)
    progress_bar.pack()
    progress_bar_inner = tk.Frame(progress_bar, height=6, bg="#3498db", width=0)
    progress_bar_inner.place(relx=0, rely=0, relheight=1)
    
    # Mensaje de estado
    estado_label = tk.Label(frame_aviso, 
                           text="Verificando conexión...",
                           font=("Arial", 9),
                           fg="#7f8c8d",
                           bg="#ffffff")
    estado_label.pack(pady=5)
    
    # Contador de tiempo
    tiempo_inicio = time.time()
    tiempo_label = tk.Label(frame_aviso, 
                           text="Tiempo esperando: 0 segundos",
                           font=("Arial", 8),
                           fg="#95a5a6",
                           bg="#ffffff")
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
            cambiar_estado("🟢 Conectado")
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
    cambiar_estado("🟡 Verificando conexión...")
    ventana.update()

    matricula = entrada.get().strip()

    if not matricula:
        messagebox.showwarning("Campo vacío", "Por favor ingresa tu matrícula.")
        entrada.focus()
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

    # 👉 SOLO confirmación, sin try/except
    mostrar_confirmacion_simple(nombre, matricula)


# --- VENTANA PRINCIPAL ---
ventana = tk.Tk()
ventana.title(f"SISTEMA DE CONTROL DE LAPTOPS - UTP | {VERSION_SISTEMA} (trabajo)")


# ✅ CONFIGURACIÓN MÁS PERMISIVA PARA LA VENTANA PRINCIPAL
ventana.attributes('-topmost', True)
ventana.state("zoomed")
ventana.overrideredirect(True)
ventana.configure(bg=COLOR_FONDO_APP)

# ✅ EVITAR GRAB_SET PARA PERMITIR MINIMIZAR
ventana.focus_force()

# Frame principal
# Sombra
sombra = tk.Frame(
    ventana,
    bg="#0a2236"
)
sombra.place(relx=0.5, rely=0.5, anchor="center", width=470, height=600)

# Tarjeta principal
main_frame = tk.Frame(
    ventana,
    bg=COLOR_TARJETA,
    bd=0,
    relief=tk.FLAT,
    width=450,
    height=580
)
main_frame.place(relx=0.5, rely=0.5, anchor="center")

# Evitar que el frame se redimensione
main_frame.pack_propagate(False)

# LOGO PRINCIPAL
logo_principal = cargar_logo("UTP.png", 160, 120)

if logo_principal:
    lbl_logo = tk.Label(main_frame, image=logo_principal, bg="#ffffff")
    lbl_logo.image = logo_principal
    lbl_logo.pack(pady=(25, 15))
else:
    tk.Label(main_frame, text="🏢", font=("Arial", 40), bg="#ffffff").pack(pady=(25, 10))
    tk.Label(main_frame, text="UTP", font=("Arial", 20, "bold"), fg="#2c3e50", bg="#ffffff").pack(pady=(0, 5))

# Títulos
tk.Label(
    main_frame,
    text="SISTEMA DE CONTROL",
    font=FUENTE_TITULO,
    fg=COLOR_TEXTO,
    bg=COLOR_TARJETA
).pack(pady=(0, 2))

tk.Label(
    main_frame,
    text="DE LAPTOPS",
    font=FUENTE_TITULO,
    fg=COLOR_TEXTO,
    bg=COLOR_TARJETA
).pack(pady=(0, 20))


# Línea decorativa
tk.Frame(main_frame, height=1, bg=COLOR_LINEA, bd=0).pack(fill=tk.X, padx=50, pady=15)

# Instrucción
tk.Label(
    main_frame,
    text="INGRESA TU MATRÍCULA",
    font=FUENTE_SUBTITULO,
    fg=COLOR_SECUNDARIO,
    bg=COLOR_TARJETA
).pack(pady=(0, 15))


# Campo de entrada
entrada_frame = tk.Frame(main_frame, bg="#ecf0f1", relief=tk.FLAT, bd=0)
entrada_frame.pack(pady=7)

entrada = RoundedEntry(
    entrada_frame,
    font=("Segoe UI", 16),
                      width=23,
                      fg="#2c3e50",
                      insertbackground="#3498db",
                      justify=tk.CENTER)
entrada.pack(pady=4, ipady=4, padx=4)
entrada.focus()
entrada.bind("<Return>", lambda event: iniciar_sesion())

# Botón
btn_entrar = RoundedButton(
    main_frame,
    text="🔐 INGRESAR",
    font=FUENTE_BOTON,
    bg=COLOR_PRIMARIO,
    fg="white",
    activebackground="#16a085",
    activeforeground="white",
    padx=35,
    pady=12,
    command=iniciar_sesion
)

btn_entrar.pack(pady=15)

# Estado de conexión a la derecha
estado_var = tk.StringVar(value="🟡 Conectando...")
estado_label = tk.Label(main_frame, 
                        textvariable=estado_var, 
                        font=("Arial", 10, "bold"),
                        fg="#27ae60", 
                        bg="#ffffff")
estado_label.pack(pady=5, anchor='e', padx=20)

# Footer (Universidad + versión)
footer_frame = tk.Frame(main_frame, bg="#ffffff")
footer_frame.pack(side=tk.BOTTOM, pady=(0, 20))


tk.Label(
    footer_frame,
    text="UNIVERSIDAD TECNOLÓGICA DE PARRAL",
    font=FUENTE_PIE,
    fg=COLOR_SECUNDARIO,
    bg=COLOR_TARJETA
).pack()


tk.Label(
    footer_frame,
    text=f"Versión {VERSION_SISTEMA}",
    font=FUENTE_PIE,
    fg=COLOR_SECUNDARIO,
    bg=COLOR_TARJETA
).pack()





ventana.protocol("WM_DELETE_WINDOW", lambda: detener_verificacion_conexion() or ventana.destroy())
# Combinación secreta para administrador: Ctrl + Alt + U
ventana.bind("<Control-Alt-u>", cerrar_sistema_admin)
ventana.bind("<Control-Alt-U>", cerrar_sistema_admin)
# Bloquear Alt + F4
ventana.bind_all("<Alt-F4>", bloquear_alt_f4)
# Evitar cierre por combinación indirecta
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