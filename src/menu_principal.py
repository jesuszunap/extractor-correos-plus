import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox


# ============================================================
# Configuración
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
EXTRACTOR_PY = BASE_DIR / "extractor_de_correos.py"
CONFIGURADOR_PY = BASE_DIR / "configurar_personas.py"
PERSONAS_JSON = BASE_DIR / "personas.json"

APP_NAME = "Extractor Correos +"


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
    """
    Ejecuta otro archivo Python usando el mismo intérprete.
    Se usa subprocess para que cada herramienta corra aparte.
    """
    subprocess.Popen(
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
        self.geometry("520x360")
        self.resizable(False, False)

        self.configurar_estilos()
        self.crear_interfaz()
        self.centrar_ventana()

    def configurar_estilos(self):
        style = ttk.Style(self)

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Titulo.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Subtitulo.TLabel", font=("Segoe UI", 10))
        style.configure("BotonGrande.TButton", font=("Segoe UI", 11), padding=10)
        style.configure("Info.TLabel", font=("Segoe UI", 9))

    def crear_interfaz(self):
        contenedor = ttk.Frame(self, padding=24)
        contenedor.pack(fill="both", expand=True)

        ttk.Label(
            contenedor,
            text=APP_NAME,
            style="Titulo.TLabel"
        ).pack(anchor="center")

        ttk.Label(
            contenedor,
            text="Sistema de exportación de correos y gestión de personas",
            style="Subtitulo.TLabel"
        ).pack(anchor="center", pady=(6, 22))

        ttk.Button(
            contenedor,
            text="Exportar correos",
            style="BotonGrande.TButton",
            command=self.abrir_extractor
        ).pack(fill="x", pady=6)

        ttk.Button(
            contenedor,
            text="Configurar personas",
            style="BotonGrande.TButton",
            command=self.abrir_configurador
        ).pack(fill="x", pady=6)

        ttk.Button(
            contenedor,
            text="Verificar archivos",
            style="BotonGrande.TButton",
            command=self.verificar_archivos
        ).pack(fill="x", pady=6)

        ttk.Button(
            contenedor,
            text="Salir",
            command=self.destroy
        ).pack(fill="x", pady=(18, 6))

        self.estado = ttk.Label(
            contenedor,
            text=self.obtener_estado_inicial(),
            style="Info.TLabel",
            foreground="#555"
        )
        self.estado.pack(anchor="center", pady=(12, 0))

    def obtener_estado_inicial(self):
        if PERSONAS_JSON.exists():
            return "personas.json detectado correctamente."
        return "Advertencia: no se encontró personas.json."

    def abrir_extractor(self):
        if not existe_archivo(EXTRACTOR_PY, "extractor_de_correos.py"):
            return

        if not existe_archivo(PERSONAS_JSON, "personas.json"):
            respuesta = messagebox.askyesno(
                "personas.json no encontrado",
                "No se encontró personas.json.\n\nEl extractor puede abrir, pero no reconocerá cargos/personas.\n\n¿Desea continuar?"
            )
            if not respuesta:
                return

        ejecutar_python(EXTRACTOR_PY)

    def abrir_configurador(self):
        if not existe_archivo(CONFIGURADOR_PY, "configurar_personas.py"):
            return

        ejecutar_python(CONFIGURADOR_PY)

    def verificar_archivos(self):
        mensajes = []

        archivos = [
            (EXTRACTOR_PY, "extractor_de_correos.py"),
            (CONFIGURADOR_PY, "configurar_personas.py"),
            (PERSONAS_JSON, "personas.json"),
        ]

        for path, nombre in archivos:
            estado = "OK" if path.exists() else "FALTA"
            mensajes.append(f"{estado}: {nombre}")

        messagebox.showinfo("Verificación", "\n".join(mensajes))
        self.estado.config(text=self.obtener_estado_inicial())

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
