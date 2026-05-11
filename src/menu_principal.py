import os
import ctypes
import logging
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import pythoncom
from extractor_de_correos import (
    LOG_PATH,
    procesar_exportacion,
    obtener_rango_dia,
    obtener_rango_mes,
)


# ============================================================
# Configuración
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
CONFIGURADOR_PY = BASE_DIR / "configurar_personas.py"
PERSONAS_JSON = BASE_DIR / "personas.json"
ICON_PATH = PROJECT_DIR / "icons" / "icono_mail.ico"

APP_NAME = "Extractor Correos +"
ANIO_ACTUAL = datetime.now().year
ANIOS_HISTORICO_SELECTOR = 3

logger = logging.getLogger("extractor_correos.gui")


# ============================================================
# Utilidades
# ============================================================

def existe_archivo(path: Path, nombre: str) -> bool:
    if not path.exists():
        messagebox.showerror(
            "Archivo no encontrado",
            f"No se encontró {nombre}:\n\n{path}"
        )
        return False
    return True


def ejecutar_python_oculto(path: Path):
    """
    Ejecuta otro archivo Python sin mostrar ventana adicional de consola.
    En Windows intenta usar pythonw.exe; si no existe, usa sys.executable ocultando ventana.
    """
    executable = Path(sys.executable)

    if executable.name.lower() == "python.exe":
        pythonw = executable.with_name("pythonw.exe")
        if pythonw.exists():
            executable = pythonw

    kwargs = {
        "cwd": str(BASE_DIR),
    }

    if sys.platform.startswith("win"):
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs["startupinfo"] = startupinfo
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

    return subprocess.Popen([str(executable), str(path)], **kwargs)


def configurar_icono_ventana(ventana):
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "ug.extractor.correos.plus"
            )
        except Exception:
            logger.exception("No se pudo establecer AppUserModelID")

    if ICON_PATH.exists():
        try:
            ventana.iconbitmap(default=str(ICON_PATH))
        except Exception:
            logger.exception("No se pudo establecer icono de ventana: %s", ICON_PATH)


# ============================================================
# Aplicación principal
# ============================================================

class MenuPrincipalApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        configurar_icono_ventana(self)
        self.geometry("820x660")
        self.minsize(820, 660)
        self.resizable(True, True)

        self.tipo_exportacion = tk.StringVar(value="recibidos")
        self.mes_completo = tk.BooleanVar(value=False)

        hoy = datetime.now()
        self.dia_var = tk.StringVar(value="")
        self.mes_var = tk.StringVar(value=str(hoy.month))
        self.anio_var = tk.IntVar(value=hoy.year)

        self.estado_var = tk.StringVar(value=self.obtener_estado_inicial())
        self.resultado_carpeta = None
        self.proceso_configurador = None
        self.exportando = False

        self.configurar_estilos()
        self.crear_interfaz()
        self.actualizar_modo_fecha()
        self.centrar_ventana()

    # --------------------------------------------------------
    # Estilos
    # --------------------------------------------------------

    def configurar_estilos(self):
        style = ttk.Style(self)

        try:
            style.theme_use("clam")
        except Exception:
            pass

        self.option_add("*Font", ("Segoe UI", 13))

        style.configure("Titulo.TLabel", font=("Segoe UI", 24, "bold"))
        style.configure("Subtitulo.TLabel", font=("Segoe UI", 14))
        style.configure("Seccion.TLabelframe.Label", font=("Segoe UI", 14, "bold"))
        style.configure("TLabel", font=("Segoe UI", 13))
        style.configure("TRadiobutton", font=("Segoe UI", 13))
        style.configure("TCheckbutton", font=("Segoe UI", 13))
        style.configure("TEntry", font=("Segoe UI", 13), padding=4)
        style.configure("TButton", font=("Segoe UI", 13), padding=8)
        style.configure("BotonGrande.TButton", font=("Segoe UI", 15, "bold"), padding=12)
        style.configure("Info.TLabel", font=("Segoe UI", 12))

    # --------------------------------------------------------
    # Interfaz
    # --------------------------------------------------------

    def crear_interfaz(self):
        contenedor = ttk.Frame(self, padding=14)
        contenedor.pack(fill="both", expand=True)

        ttk.Frame(contenedor).pack(pady=(2, 2))

        # Tipo de exportación
        frame_tipo = ttk.LabelFrame(contenedor, text="Tipo de exportación", padding=14, style="Seccion.TLabelframe")
        frame_tipo.pack(fill="x", pady=(0, 14))

        ttk.Radiobutton(
            frame_tipo,
            text="Correos recibidos",
            variable=self.tipo_exportacion,
            value="recibidos"
        ).grid(row=0, column=0, sticky="w", padx=(0, 36))

        ttk.Radiobutton(
            frame_tipo,
            text="Correos enviados",
            variable=self.tipo_exportacion,
            value="enviados"
        ).grid(row=0, column=1, sticky="w")

        # Fecha
        frame_fecha = ttk.LabelFrame(contenedor, text="Fecha de exportación", padding=14, style="Seccion.TLabelframe")
        frame_fecha.pack(fill="x", pady=(0, 14))

        self.check_mes = ttk.Checkbutton(
            frame_fecha,
            text="Exportar mes completo",
            variable=self.mes_completo,
            command=self.actualizar_modo_fecha
        )

        self.tipo_exportacion.set("recibidos")
        self.check_mes.grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 12))

        self.dia_label = ttk.Label(frame_fecha, text="Día:")
        self.dia_label.grid(row=1, column=0, sticky="w")
        self.dia_entry = ttk.Entry(frame_fecha, textvariable=self.dia_var, width=8)
        self.dia_entry.grid(row=1, column=1, sticky="w", padx=(8, 22))

        ttk.Label(frame_fecha, text="Mes:").grid(row=1, column=2, sticky="w")

        meses_display = [
            "(1) Enero",
            "(2) Febrero",
            "(3) Marzo",
            "(4) Abril",
            "(5) Mayo",
            "(6) Junio",
            "(7) Julio",
            "(8) Agosto",
            "(9) Septiembre",
            "(10) Octubre",
            "(11) Noviembre",
            "(12) Diciembre",
        ]

        self.mes_combo = ttk.Combobox(
            frame_fecha,
            values=meses_display,
            state="readonly",
            width=18
        )

        self.mes_combo.grid(row=1, column=3, sticky="w", padx=(8, 22))

        # Seleccionar mes actual automáticamente
        self.mes_combo.current(datetime.now().month - 1)

        ttk.Label(frame_fecha, text="Año:").grid(row=1, column=4, sticky="w")
        self.anios_display = [
            str(a)
            for a in range(ANIO_ACTUAL, ANIO_ACTUAL - ANIOS_HISTORICO_SELECTOR - 1, -1)
        ]

        self.anio_combo = ttk.Combobox(
            frame_fecha,
            values=self.anios_display,
            state="readonly",
            width=10
        )
        self.anio_combo.grid(row=1, column=5, sticky="w", padx=(8, 0))
        self.anio_combo.set(str(ANIO_ACTUAL))

        frame_fecha.columnconfigure(5, weight=1)

        # Botón principal
        self.boton_exportar = ttk.Button(
            contenedor,
            text="Exportar correos",
            style="BotonGrande.TButton",
            command=self.iniciar_exportacion
        )
        self.boton_exportar.pack(fill="x", pady=(4, 14))

        # Progreso
        frame_progreso = ttk.LabelFrame(contenedor, text="Progreso", padding=14, style="Seccion.TLabelframe")
        frame_progreso.pack(fill="x", pady=(0, 14))

        self.progress = ttk.Progressbar(frame_progreso, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(0, 10))

        self.label_estado_progreso = ttk.Label(
            frame_progreso,
            textvariable=self.estado_var,
            style="Info.TLabel",
            foreground="#555",
            wraplength=690
        )
        self.label_estado_progreso.pack(anchor="w")

        self.boton_abrir_carpeta = ttk.Button(
            frame_progreso,
            text="Abrir carpeta de exportación",
            command=self.abrir_carpeta_resultado
        )
        self.boton_abrir_carpeta.pack(fill="x", pady=(12, 0))
        self.boton_abrir_carpeta.pack_forget()

        # Herramientas secundarias
        frame_herramientas = ttk.LabelFrame(contenedor, text="Herramientas", padding=14, style="Seccion.TLabelframe")
        frame_herramientas.pack(fill="x", pady=(0, 14))

        ttk.Button(
            frame_herramientas,
            text="Configurar personas",
            command=self.abrir_configurador
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ttk.Button(
            frame_herramientas,
            text="Verificar archivos",
            command=self.verificar_archivos
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        frame_herramientas.columnconfigure(0, weight=1)
        frame_herramientas.columnconfigure(1, weight=1)

    # --------------------------------------------------------
    # Estado y validaciones
    # --------------------------------------------------------

    def obtener_estado_inicial(self):
        if PERSONAS_JSON.exists():
            return "Listo. personas.json detectado correctamente."
        return "Advertencia: no se encontró personas.json."

    def actualizar_modo_fecha(self):
        usar_mes = self.mes_completo.get()

        if usar_mes:
            self.dia_entry.config(state="disabled")
        else:
            self.dia_entry.config(state="normal")

    def validar_entero(self, valor, nombre, minimo, maximo):
        try:
            numero = int(valor)
        except ValueError:
            raise ValueError(f"{nombre} debe ser un número.")

        if numero < minimo or numero > maximo:
            raise ValueError(f"{nombre} debe estar entre {minimo} y {maximo}.")

        return numero

    def obtener_parametros_exportacion(self):
        tipo = self.tipo_exportacion.get().strip() or "recibidos"

        if tipo not in ["recibidos", "enviados"]:
            raise ValueError("Seleccione si desea exportar recibidos o enviados.")

        mes_texto = self.mes_combo.get()

        if not mes_texto:
            raise ValueError("Seleccione un mes.")

        mes = int(mes_texto.split(")")[0].replace("(", ""))

        anio = int(self.anio_combo.get())

        if self.mes_completo.get():
            fecha_inicio, fecha_fin, etiqueta, _mes_name, _anio = obtener_rango_mes(mes, anio)
            return tipo, fecha_inicio, fecha_fin, etiqueta

        dia = self.validar_entero(self.dia_var.get(), "Día", 1, 31)

        try:
            fecha_inicio, fecha_fin, etiqueta, _mes_name, _anio = obtener_rango_dia(dia, mes, anio)
        except ValueError:
            raise ValueError("La fecha ingresada no es válida.")

        return tipo, fecha_inicio, fecha_fin, etiqueta

    def set_exportando(self, exportando):
        self.exportando = exportando
        estado = "disabled" if exportando else "normal"

        self.boton_exportar.config(state=estado)
        self.check_mes.config(state=estado)
        self.dia_entry.config(state=estado)
        self.mes_combo.config(state="disabled" if exportando else "readonly")
        self.anio_combo.config(state="disabled" if exportando else "readonly")

        if not exportando:
            self.actualizar_modo_fecha()

        if exportando:
            self.progress["value"] = 0
            self.boton_abrir_carpeta.pack_forget()
            self.resultado_carpeta = None
        else:
            if self.progress["value"] < 100:
                self.progress["value"] = 0

    # --------------------------------------------------------
    # Exportación
    # --------------------------------------------------------

    def iniciar_exportacion(self):
        if self.exportando:
            return

        logger.info("Usuario inicio exportacion desde GUI")

        if not existe_archivo(PERSONAS_JSON, "personas.json"):
            respuesta = messagebox.askyesno(
                "personas.json no encontrado",
                "No se encontró personas.json.\n\nEl extractor puede continuar, pero no reconocerá cargos/personas.\n\n¿Desea continuar?"
            )
            if not respuesta:
                return

        try:
            tipo, fecha_inicio, fecha_fin, etiqueta = self.obtener_parametros_exportacion()
        except ValueError as e:
            messagebox.showerror("Datos inválidos", str(e))
            return

        self.set_exportando(True)
        self.estado_var.set("Preparando exportación...")

        hilo = threading.Thread(
            target=self.ejecutar_exportacion_en_hilo,
            args=(tipo, fecha_inicio, fecha_fin, etiqueta),
            daemon=True
        )
        hilo.start()

    def ejecutar_exportacion_en_hilo(self, tipo, fecha_inicio, fecha_fin, etiqueta):
        pythoncom.CoInitialize()

        try:
            resultado = procesar_exportacion(
                tipo_exportacion=tipo,
                fecha_inicio=fecha_inicio,
                fecha_fin=fecha_fin,
                etiqueta_fecha=etiqueta,
                progress_callback=self.actualizar_progreso_desde_hilo
            )

            self.after(0, lambda: self.finalizar_exportacion(resultado))

        except Exception as e:
            self.after(0, lambda error=e: self.exportacion_con_error(error))

        finally:
            pythoncom.CoUninitialize()

    def actualizar_progreso_desde_hilo(self, progreso):
        def aplicar_progreso():
            if isinstance(progreso, dict):
                mensaje = progreso.get("mensaje", "")
                porcentaje = progreso.get("porcentaje")

                if porcentaje is not None:
                    self.progress["value"] = max(0, min(100, int(porcentaje)))

                self.estado_var.set(str(mensaje))
                return

            self.estado_var.set(str(progreso))

        self.after(0, aplicar_progreso)

    def finalizar_exportacion(self, resultado):
        self.set_exportando(False)
        self.estado_var.set(resultado.mensaje)
        logger.info(
            "Exportacion finalizada en GUI. Exito=%s, cantidad=%s, carpeta=%s",
            resultado.exito,
            resultado.cantidad,
            resultado.carpeta
        )

        if resultado.exito and resultado.carpeta:
            self.resultado_carpeta = resultado.carpeta
            self.boton_abrir_carpeta.pack(fill="x", pady=(12, 0))
        else:
            self.resultado_carpeta = None
            self.boton_abrir_carpeta.pack_forget()
            messagebox.showwarning("Sin registros", resultado.mensaje)

    def exportacion_con_error(self, error):
        self.set_exportando(False)
        self.estado_var.set("Ocurrió un error durante la exportación.")
        logger.error(
            "Error durante la exportacion desde GUI",
            exc_info=(type(error), error, error.__traceback__)
        )
        mensaje = str(error)

        if "-2146959355" in mensaje or "-2079063787" in mensaje or "0x84140115" in mensaje:
            mensaje = (
                "Outlook o Word no respondio a tiempo.\n\n"
                "Revise si Outlook tiene una ventana abierta preguntando si desea conectar, "
                "pulse Conectar y vuelva a intentar.\n\n"
                "Si la conexion esta lenta, espere unos minutos y ejecute nuevamente."
            )
        elif "-2147418111" in mensaje or "-2147417846" in mensaje:
            mensaje = (
                "Outlook o Word esta ocupado.\n\n"
                "Cierre cuadros de dialogo abiertos de Outlook/Word y vuelva a intentar."
            )

        self.estado_var.set("Ocurrio un error durante la exportacion.")
        messagebox.showerror("Error", f"{mensaje}\n\nLog:\n{LOG_PATH}")

    def abrir_carpeta_resultado(self):
        if not self.resultado_carpeta:
            return

        try:
            os.startfile(self.resultado_carpeta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n\n{e}")

    # --------------------------------------------------------
    # Herramientas
    # --------------------------------------------------------

    def abrir_configurador(self):
        if self.proceso_configurador and self.proceso_configurador.poll() is None:
            messagebox.showinfo(
                "Configurador abierto",
                "El configurador de personas ya está abierto. Cierre esa ventana antes de abrir otra."
            )
            return

        if not existe_archivo(CONFIGURADOR_PY, "configurar_personas.py"):
            return

        self.proceso_configurador = ejecutar_python_oculto(CONFIGURADOR_PY)

    def verificar_archivos(self):
        archivos = [
            (CONFIGURADOR_PY, "configurar_personas.py"),
            (PERSONAS_JSON, "personas.json"),
        ]

        mensajes = []
        for path, nombre in archivos:
            estado = "OK" if path.exists() else "FALTA"
            mensajes.append(f"{estado}: {nombre}")

        messagebox.showinfo("Verificación", "\n".join(mensajes))
        self.estado_var.set(self.obtener_estado_inicial())

    def centrar_ventana(self):
        self.update_idletasks()
        ancho = self.winfo_width()
        alto = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.winfo_screenheight() // 2) - (alto // 2)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")


if __name__ == "__main__":
    app = MenuPrincipalApp()
    app.mainloop()
