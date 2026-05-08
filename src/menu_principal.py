import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

from extractor_de_correos import (
    procesar_exportacion,
    obtener_rango_dia,
    obtener_rango_mes,
)


# ============================================================
# Configuración
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
CONFIGURADOR_PY = BASE_DIR / "configurar_personas.py"
PERSONAS_JSON = BASE_DIR / "personas.json"

APP_NAME = "Extractor Correos +"
MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]


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


def ejecutar_python(path: Path):
    return subprocess.Popen(
        [sys.executable, str(path)],
        cwd=str(BASE_DIR),
        creationflags=subprocess.CREATE_NEW_CONSOLE if sys.platform.startswith("win") else 0
    )


# ============================================================
# Aplicación principal
# ============================================================

class MenuPrincipalApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        self.geometry("700x620")
        self.minsize(700, 620)
        self.resizable(False, False)

        self.tipo_exportacion = tk.StringVar(value="recibidos")
        self.mes_completo = tk.BooleanVar(value=False)

        hoy = datetime.now()
        self.dia_var = tk.StringVar(value=str(hoy.day))
        self.mes_var = tk.StringVar(value=str(hoy.month))
        self.anio_var = tk.StringVar(value=str(hoy.year))

        self.mes_combo_var = tk.StringVar(value=MESES[hoy.month - 1])
        self.anio_mes_var = tk.StringVar(value=str(hoy.year))

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

        style.configure("Titulo.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Subtitulo.TLabel", font=("Segoe UI", 10))
        style.configure("Seccion.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("BotonGrande.TButton", font=("Segoe UI", 11, "bold"), padding=10)
        style.configure("Info.TLabel", font=("Segoe UI", 9))
        style.configure("Resultado.TLabel", font=("Segoe UI", 10, "bold"))

    # --------------------------------------------------------
    # Interfaz
    # --------------------------------------------------------

    def crear_interfaz(self):
        contenedor = ttk.Frame(self, padding=20)
        contenedor.pack(fill="both", expand=True)

        ttk.Label(contenedor, text=APP_NAME, style="Titulo.TLabel").pack(anchor="center")
        ttk.Label(
            contenedor,
            text="Exportación de correos institucionales y gestión de personas",
            style="Subtitulo.TLabel"
        ).pack(anchor="center", pady=(6, 16))

        # Tipo de exportación
        frame_tipo = ttk.LabelFrame(contenedor, text="Tipo de exportación", padding=12, style="Seccion.TLabelframe")
        frame_tipo.pack(fill="x", pady=(0, 12))

        ttk.Radiobutton(
            frame_tipo,
            text="Correos recibidos",
            variable=self.tipo_exportacion,
            value="recibidos"
        ).grid(row=0, column=0, sticky="w", padx=(0, 24))

        ttk.Radiobutton(
            frame_tipo,
            text="Correos enviados",
            variable=self.tipo_exportacion,
            value="enviados"
        ).grid(row=0, column=1, sticky="w")

        # Fecha
        frame_fecha = ttk.LabelFrame(contenedor, text="Fecha de exportación", padding=12, style="Seccion.TLabelframe")
        frame_fecha.pack(fill="x", pady=(0, 12))

        self.check_mes = ttk.Checkbutton(
            frame_fecha,
            text="Exportar mes completo",
            variable=self.mes_completo,
            command=self.actualizar_modo_fecha
        )
        self.check_mes.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        ttk.Label(frame_fecha, text="Día:").grid(row=1, column=0, sticky="w")
        self.dia_entry = ttk.Entry(frame_fecha, textvariable=self.dia_var, width=8)
        self.dia_entry.grid(row=1, column=1, sticky="w", padx=(6, 16))

        ttk.Label(frame_fecha, text="Mes:").grid(row=1, column=2, sticky="w")
        self.mes_entry = ttk.Entry(frame_fecha, textvariable=self.mes_var, width=8)
        self.mes_entry.grid(row=1, column=3, sticky="w", padx=(6, 16))

        ttk.Label(frame_fecha, text="Año:").grid(row=1, column=4, sticky="w")
        self.anio_entry = ttk.Entry(frame_fecha, textvariable=self.anio_var, width=10)
        self.anio_entry.grid(row=1, column=5, sticky="w", padx=(6, 0))

        ttk.Label(frame_fecha, text="Mes completo:").grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.mes_combo = ttk.Combobox(
            frame_fecha,
            textvariable=self.mes_combo_var,
            values=MESES,
            state="readonly",
            width=16
        )
        self.mes_combo.grid(row=2, column=1, columnspan=2, sticky="w", padx=(6, 16), pady=(12, 0))

        ttk.Label(frame_fecha, text="Año:").grid(row=2, column=3, sticky="w", pady=(12, 0))
        self.anio_mes_entry = ttk.Entry(frame_fecha, textvariable=self.anio_mes_var, width=10)
        self.anio_mes_entry.grid(row=2, column=4, sticky="w", padx=(6, 0), pady=(12, 0))

        # Botón principal
        self.boton_exportar = ttk.Button(
            contenedor,
            text="Exportar correos",
            style="BotonGrande.TButton",
            command=self.iniciar_exportacion
        )
        self.boton_exportar.pack(fill="x", pady=(4, 12))

        # Progreso
        frame_progreso = ttk.LabelFrame(contenedor, text="Progreso", padding=12, style="Seccion.TLabelframe")
        frame_progreso.pack(fill="x", pady=(0, 12))

        self.progress = ttk.Progressbar(frame_progreso, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 8))

        self.label_estado_progreso = ttk.Label(
            frame_progreso,
            textvariable=self.estado_var,
            style="Info.TLabel",
            foreground="#555",
            wraplength=630
        )
        self.label_estado_progreso.pack(anchor="w")

        self.boton_abrir_carpeta = ttk.Button(
            frame_progreso,
            text="Abrir carpeta de exportación",
            command=self.abrir_carpeta_resultado
        )
        self.boton_abrir_carpeta.pack(fill="x", pady=(10, 0))
        self.boton_abrir_carpeta.pack_forget()

        # Herramientas secundarias
        frame_herramientas = ttk.LabelFrame(contenedor, text="Herramientas", padding=12, style="Seccion.TLabelframe")
        frame_herramientas.pack(fill="x", pady=(0, 12))

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

        ttk.Button(contenedor, text="Salir", command=self.destroy).pack(fill="x")

    # --------------------------------------------------------
    # Estado y validaciones
    # --------------------------------------------------------

    def obtener_estado_inicial(self):
        if PERSONAS_JSON.exists():
            return "Listo. personas.json detectado correctamente."
        return "Advertencia: no se encontró personas.json."

    def actualizar_modo_fecha(self):
        usar_mes = self.mes_completo.get()

        estado_fecha = "disabled" if usar_mes else "normal"
        estado_mes = "readonly" if usar_mes else "disabled"
        estado_anio_mes = "normal" if usar_mes else "disabled"

        self.dia_entry.config(state=estado_fecha)
        self.mes_entry.config(state=estado_fecha)
        self.anio_entry.config(state=estado_fecha)
        self.mes_combo.config(state=estado_mes)
        self.anio_mes_entry.config(state=estado_anio_mes)

    def validar_entero(self, valor, nombre, minimo, maximo):
        try:
            numero = int(valor)
        except ValueError:
            raise ValueError(f"{nombre} debe ser un número.")

        if numero < minimo or numero > maximo:
            raise ValueError(f"{nombre} debe estar entre {minimo} y {maximo}.")

        return numero

    def obtener_parametros_exportacion(self):
        tipo = self.tipo_exportacion.get()

        if tipo not in ["recibidos", "enviados"]:
            raise ValueError("Seleccione si desea exportar recibidos o enviados.")

        if self.mes_completo.get():
            mes = MESES.index(self.mes_combo_var.get()) + 1
            anio = self.validar_entero(self.anio_mes_var.get(), "Año", 1900, 2100)
            fecha_inicio, fecha_fin, etiqueta, _mes_name, _anio = obtener_rango_mes(mes, anio)
            return tipo, fecha_inicio, fecha_fin, etiqueta

        dia = self.validar_entero(self.dia_var.get(), "Día", 1, 31)
        mes = self.validar_entero(self.mes_var.get(), "Mes", 1, 12)
        anio = self.validar_entero(self.anio_var.get(), "Año", 1900, 2100)

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
        self.actualizar_modo_fecha()

        if exportando:
            self.progress.start(12)
            self.boton_abrir_carpeta.pack_forget()
            self.resultado_carpeta = None
        else:
            self.progress.stop()

    # --------------------------------------------------------
    # Exportación
    # --------------------------------------------------------

    def iniciar_exportacion(self):
        if self.exportando:
            return

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
            self.after(0, lambda: self.exportacion_con_error(e))

    def actualizar_progreso_desde_hilo(self, mensaje):
        self.after(0, lambda: self.estado_var.set(str(mensaje)))

    def finalizar_exportacion(self, resultado):
        self.set_exportando(False)
        self.estado_var.set(resultado.mensaje)

        if resultado.exito and resultado.carpeta:
            self.resultado_carpeta = resultado.carpeta
            self.boton_abrir_carpeta.pack(fill="x", pady=(10, 0))
            messagebox.showinfo("Exportación completada", resultado.mensaje)
        else:
            self.resultado_carpeta = None
            self.boton_abrir_carpeta.pack_forget()
            messagebox.showwarning("Sin registros", resultado.mensaje)

    def exportacion_con_error(self, error):
        self.set_exportando(False)
        self.estado_var.set("Ocurrió un error durante la exportación.")
        messagebox.showerror("Error", str(error))

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

        self.proceso_configurador = ejecutar_python(CONFIGURADOR_PY)

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
