import json
import calendar
import shutil
import uuid
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from config_rutas import obtener_personas_json


# ============================================================
# Configuración
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
PERSONAS_JSON = obtener_personas_json()
BACKUP_DIR = PERSONAS_JSON.parent / "backups_personas"

COLUMNAS = [
    "registro_id",
    "nombre_completo",
    "nombre_corto",
    "titulo_abreviado",
    "cargo",
    "dependencia",
    "email",
    "activo",
    "vigente_desde",
    "vigente_hasta",
    "area_base",
    "fuente",
]

COLUMNAS_VISIBLES = {
    "registro_id": "ID",
    "nombre_completo": "Nombre completo",
    "nombre_corto": "Nombre corto",
    "titulo_abreviado": "Título",
    "cargo": "Cargo",
    "dependencia": "Dependencia",
    "email": "E-mail",
    "activo": "Activo",
    "vigente_desde": "Desde",
    "vigente_hasta": "Hasta",
    "area_base": "Área base",
    "fuente": "Fuente",
}

CAMPOS_FORMULARIO = [
    ("nombre_completo", "Nombre completo"),
    ("nombre_corto", "Nombre corto"),
    ("titulo", "Título completo"),
    ("titulo_abreviado", "Título abreviado"),
    ("cargo", "Cargo"),
    ("dependencia", "Dependencia"),
    ("email", "E-mail"),
    ("area_base", "Área base"),
    ("vigente_desde", "Vigente desde (YYYY-MM-DD)"),
    ("vigente_hasta", "Vigente hasta (YYYY-MM-DD)"),
    ("fuente", "Fuente"),
]

CAMPOS_FECHA = {"vigente_desde", "vigente_hasta"}
MESES_CALENDARIO = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
]
DIAS_CALENDARIO = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]


# ============================================================
# Utilidades
# ============================================================

def normalizar_texto(texto):
    return str(texto or "").strip()


def bool_a_texto(valor):
    if valor is True:
        return "Sí"
    if valor is False:
        return "No"
    return ""


def texto_a_bool(valor):
    valor = str(valor or "").strip().lower()
    if valor in ["sí", "si", "s", "true", "1", "activo"]:
        return True
    if valor in ["no", "n", "false", "0", "inactivo"]:
        return False
    return None


def validar_fecha_o_vacio(valor):
    valor = normalizar_texto(valor)
    if not valor:
        return True
    try:
        datetime.strptime(valor, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def validar_email_o_vacio(valor):
    valor = normalizar_texto(valor)
    if not valor:
        return True

    if " " in valor or valor.count("@") != 1:
        return False

    usuario, dominio = valor.split("@", 1)
    return bool(usuario.strip() and dominio.strip())


def generar_id():
    return f"manual_{uuid.uuid4().hex[:8]}"


def calcular_nombre_corto(nombre_completo):
    partes = normalizar_texto(nombre_completo).split()
    if len(partes) >= 3:
        return f"{partes[0]} {partes[-2]}"
    if len(partes) >= 2:
        return f"{partes[0]} {partes[-1]}"
    return normalizar_texto(nombre_completo)


def crear_backup():
    if not PERSONAS_JSON.exists():
        return None

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    limpiar_backups_antiguos()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup_path = BACKUP_DIR / f"personas_backup_{timestamp}.json"
    contador = 2
    while backup_path.exists():
        backup_path = BACKUP_DIR / f"personas_backup_{timestamp}_{contador}.json"
        contador += 1

    shutil.copy2(PERSONAS_JSON, backup_path)
    return backup_path


def limpiar_backups_antiguos(dias=90):
    if not BACKUP_DIR.exists():
        return 0

    limite = datetime.now() - timedelta(days=dias)
    eliminados = 0

    for backup in BACKUP_DIR.glob("personas_backup_*.json"):
        fecha = obtener_fecha_backup(backup)
        if fecha is None:
            try:
                fecha = datetime.fromtimestamp(backup.stat().st_mtime)
            except OSError:
                continue

        if fecha >= limite:
            continue

        try:
            backup.unlink()
            eliminados += 1
        except OSError:
            pass

    return eliminados


def obtener_fecha_backup(path):
    nombre = path.stem.replace("personas_backup_", "", 1)
    partes = nombre.split("_")
    if len(partes) >= 2:
        nombre = "_".join(partes[:2])

    try:
        return datetime.strptime(nombre, "%Y-%m-%d_%H-%M-%S")
    except ValueError:
        return None


def listar_backups():
    if not BACKUP_DIR.exists():
        return []

    backups = [p for p in BACKUP_DIR.glob("personas_backup_*.json") if p.is_file()]
    backups.sort(key=lambda p: obtener_fecha_backup(p) or datetime.min, reverse=True)
    return backups


def formatear_fecha_backup(path):
    fecha = obtener_fecha_backup(path)
    if fecha:
        return fecha.strftime("%Y-%m-%d %H:%M:%S")
    return "Fecha no reconocida"


# ============================================================
# Carga y guardado JSON
# ============================================================

def cargar_json():
    if not PERSONAS_JSON.exists():
        data = {
            "metadata": {
                "descripcion": "Personas reconocidas para extractor-correos-plus.",
                "total_registros": 0,
                "usa_vigencias": True,
            },
            "personas": []
        }
        return data

    with open(PERSONAS_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        data = {
            "metadata": {
                "descripcion": "Personas reconocidas para extractor-correos-plus.",
                "usa_vigencias": True,
            },
            "personas": data
        }

    if "metadata" not in data:
        data["metadata"] = {}
    if "personas" not in data or not isinstance(data["personas"], list):
        data["personas"] = []

    for persona in data["personas"]:
        persona.setdefault("registro_id", generar_id())
        persona.setdefault("nombre_completo", "")
        persona.setdefault("nombre_corto", calcular_nombre_corto(persona.get("nombre_completo", "")))
        persona.setdefault("titulo", "")
        persona.setdefault("titulo_abreviado", "")
        persona.setdefault("cargo", "")
        persona.setdefault("dependencia", "")
        persona.setdefault("email", "")
        persona.setdefault("activo", True)
        persona.setdefault("vigente_desde", None)
        persona.setdefault("vigente_hasta", None)
        persona.setdefault("area_base", "")
        persona.setdefault("fuente", "manual")

    return data


def guardar_json(data):
    data["metadata"]["total_registros"] = len(data.get("personas", []))
    data["metadata"]["usa_vigencias"] = True
    data["metadata"]["ultima_actualizacion_manual"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(PERSONAS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ============================================================
# Ventana de formulario
# ============================================================

class CalendarioDialog(tk.Toplevel):
    def __init__(self, parent, fecha_actual=None):
        super().__init__(parent)
        self.title("Seleccionar fecha")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        try:
            ttk.Style(self).configure("DiaSeleccionado.TButton", font=("Segoe UI", 9, "bold"))
        except Exception:
            pass

        self.resultado = None
        fecha_base = self.obtener_fecha_base(fecha_actual)
        self.anio = fecha_base.year
        self.mes = fecha_base.month
        self.dia_seleccionado = fecha_base.day

        contenedor = ttk.Frame(self, padding=12)
        contenedor.grid(row=0, column=0, sticky="nsew")

        cabecera = ttk.Frame(contenedor)
        cabecera.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        ttk.Button(cabecera, text="<", width=4, command=self.mes_anterior).grid(row=0, column=0)
        self.label_mes = ttk.Label(cabecera, width=22, anchor="center")
        self.label_mes.grid(row=0, column=1, padx=8)
        ttk.Button(cabecera, text=">", width=4, command=self.mes_siguiente).grid(row=0, column=2)

        self.frame_dias = ttk.Frame(contenedor)
        self.frame_dias.grid(row=1, column=0, sticky="nsew")

        botones = ttk.Frame(contenedor)
        botones.grid(row=2, column=0, sticky="e", pady=(10, 0))
        ttk.Button(botones, text="Hoy", command=self.ir_a_hoy).grid(row=0, column=0, padx=4)
        ttk.Button(botones, text="Cancelar", command=self.destroy).grid(row=0, column=1, padx=4)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())
        self.dibujar_calendario()

        self.update_idletasks()
        self.centrar(parent)

    def obtener_fecha_base(self, fecha_actual):
        if validar_fecha_o_vacio(fecha_actual) and normalizar_texto(fecha_actual):
            return datetime.strptime(fecha_actual, "%Y-%m-%d")
        return datetime.now()

    def centrar(self, parent):
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def mes_anterior(self):
        if self.mes == 1:
            self.mes = 12
            self.anio -= 1
        else:
            self.mes -= 1
        self.dibujar_calendario()

    def mes_siguiente(self):
        if self.mes == 12:
            self.mes = 1
            self.anio += 1
        else:
            self.mes += 1
        self.dibujar_calendario()

    def ir_a_hoy(self):
        hoy = datetime.now()
        self.anio = hoy.year
        self.mes = hoy.month
        self.dia_seleccionado = hoy.day
        self.dibujar_calendario()

    def dibujar_calendario(self):
        for widget in self.frame_dias.winfo_children():
            widget.destroy()

        self.label_mes.config(text=f"{MESES_CALENDARIO[self.mes - 1]} {self.anio}")

        for col, dia in enumerate(DIAS_CALENDARIO):
            ttk.Label(self.frame_dias, text=dia, anchor="center", width=5).grid(row=0, column=col, padx=1, pady=1)

        semanas = calendar.Calendar(firstweekday=0).monthdayscalendar(self.anio, self.mes)
        for fila, semana in enumerate(semanas, start=1):
            for col, dia in enumerate(semana):
                if dia == 0:
                    ttk.Label(self.frame_dias, text="", width=5).grid(row=fila, column=col, padx=1, pady=1)
                    continue

                estilo = "TButton"
                if dia == self.dia_seleccionado:
                    estilo = "DiaSeleccionado.TButton"

                ttk.Button(
                    self.frame_dias,
                    text=str(dia),
                    width=5,
                    style=estilo,
                    command=lambda d=dia: self.seleccionar_dia(d),
                ).grid(row=fila, column=col, padx=1, pady=1)

    def seleccionar_dia(self, dia):
        self.resultado = f"{self.anio:04d}-{self.mes:02d}-{dia:02d}"
        self.destroy()


class RestaurarCopiaDialog(tk.Toplevel):
    def __init__(self, parent, backups):
        super().__init__(parent)
        self.title("Restaurar copia de seguridad")
        self.geometry("780x360")
        self.minsize(680, 320)
        self.transient(parent)
        self.grab_set()

        self.backups = backups
        self.resultado = None

        contenedor = ttk.Frame(self, padding=12)
        contenedor.pack(fill="both", expand=True)

        ttk.Label(
            contenedor,
            text="Seleccione la copia de seguridad que desea restaurar:"
        ).pack(anchor="w", pady=(0, 8))

        frame_tabla = ttk.Frame(contenedor)
        frame_tabla.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            frame_tabla,
            columns=("fecha", "archivo", "ruta"),
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("fecha", text="Fecha y hora")
        self.tree.heading("archivo", text="Archivo")
        self.tree.heading("ruta", text="Ruta")
        self.tree.column("fecha", width=160, anchor="w")
        self.tree.column("archivo", width=260, anchor="w")
        self.tree.column("ruta", width=320, anchor="w")

        scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")

        frame_tabla.rowconfigure(0, weight=1)
        frame_tabla.columnconfigure(0, weight=1)

        for idx, backup in enumerate(self.backups):
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(formatear_fecha_backup(backup), backup.name, str(backup)),
            )

        botones = ttk.Frame(contenedor)
        botones.pack(fill="x", pady=(10, 0))
        ttk.Button(botones, text="Restaurar", command=self.restaurar).pack(side="right", padx=4)
        ttk.Button(botones, text="Cancelar", command=self.destroy).pack(side="right", padx=4)

        self.tree.bind("<Double-1>", lambda _event: self.restaurar())
        self.bind("<Return>", lambda _event: self.restaurar())
        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        if self.backups:
            self.tree.selection_set("0")
            self.tree.focus("0")

        self.update_idletasks()
        self.centrar(parent)

    def centrar(self, parent):
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def restaurar(self):
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Sin seleccion", "Seleccione una copia de seguridad primero.")
            return

        self.resultado = self.backups[int(seleccion[0])]
        self.destroy()


class PersonaDialog(tk.Toplevel):
    def __init__(self, parent, titulo, persona=None):
        super().__init__(parent)
        self.title(titulo)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.resultado = None
        self.persona_original = persona or {}
        self.variables = {}

        contenedor = ttk.Frame(self, padding=12)
        contenedor.grid(row=0, column=0, sticky="nsew")

        for i, (campo, etiqueta) in enumerate(CAMPOS_FORMULARIO):
            ttk.Label(contenedor, text=etiqueta).grid(row=i, column=0, sticky="w", pady=3)
            var = tk.StringVar(value="" if self.persona_original.get(campo) is None else str(self.persona_original.get(campo, "")))
            self.variables[campo] = var

            if campo in CAMPOS_FECHA:
                frame_fecha = ttk.Frame(contenedor)
                frame_fecha.grid(row=i, column=1, sticky="ew", pady=3)
                frame_fecha.columnconfigure(0, weight=1)

                entrada = ttk.Entry(frame_fecha, textvariable=var, width=58, state="readonly")
                entrada.grid(row=0, column=0, sticky="ew")
                ttk.Button(
                    frame_fecha,
                    text="Calendario",
                    command=lambda c=campo: self.abrir_calendario(c),
                ).grid(row=0, column=1, padx=(6, 0))
                ttk.Button(
                    frame_fecha,
                    text="Limpiar",
                    command=lambda v=var: v.set(""),
                ).grid(row=0, column=2, padx=(6, 0))
            else:
                entrada = ttk.Entry(contenedor, textvariable=var, width=70)
                entrada.grid(row=i, column=1, sticky="ew", pady=3)

        self.activo_var = tk.BooleanVar(value=self.persona_original.get("activo", True) is True)
        ttk.Checkbutton(contenedor, text="Activo", variable=self.activo_var).grid(row=len(CAMPOS_FORMULARIO), column=1, sticky="w", pady=8)

        nota = (
            "Nota: las fechas son opcionales. Si se dejan vacías, el registro se considera sin límite histórico."
        )
        ttk.Label(contenedor, text=nota, foreground="#555").grid(row=len(CAMPOS_FORMULARIO) + 1, column=0, columnspan=2, sticky="w", pady=(4, 10))

        botones = ttk.Frame(contenedor)
        botones.grid(row=len(CAMPOS_FORMULARIO) + 2, column=0, columnspan=2, sticky="e")

        ttk.Button(botones, text="Guardar", command=self.guardar).grid(row=0, column=0, padx=4)
        ttk.Button(botones, text="Cancelar", command=self.cerrar).grid(row=0, column=1, padx=4)

        self.bind("<Return>", lambda _event: self.guardar())
        self.bind("<Escape>", lambda _event: self.cerrar())
        self.protocol("WM_DELETE_WINDOW", self.cerrar)

        self.update_idletasks()
        self.centrar(parent)

    def centrar(self, parent):
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def abrir_calendario(self, campo):
        dialog = CalendarioDialog(self, self.variables[campo].get())
        self.wait_window(dialog)

        if dialog.resultado:
            self.variables[campo].set(dialog.resultado)

    def construir_persona_desde_formulario(self, completar_defaults=True):
        persona = dict(self.persona_original)

        for campo, var in self.variables.items():
            valor = normalizar_texto(var.get())
            persona[campo] = valor if valor else None

        persona["activo"] = self.activo_var.get()
        if completar_defaults:
            persona["registro_id"] = persona.get("registro_id") or generar_id()
            persona["fuente"] = persona.get("fuente") or "manual"
        return persona

    def normalizar_persona_para_comparar(self, persona):
        normalizada = dict(persona)

        for campo, _etiqueta in CAMPOS_FORMULARIO:
            valor = normalizar_texto(normalizada.get(campo))
            normalizada[campo] = valor if valor else None

        normalizada["activo"] = normalizada.get("activo", True) is True
        return normalizada

    def formulario_tiene_cambios(self):
        actual = self.construir_persona_desde_formulario(completar_defaults=False)
        original = self.normalizar_persona_para_comparar(self.persona_original)
        return actual != original

    def cerrar(self):
        if not self.formulario_tiene_cambios():
            self.destroy()
            return

        respuesta = messagebox.askyesnocancel(
            "Cambios sin guardar",
            "Se detectaron cambios en esta persona.\n\nDesea guardarlos antes de cerrar?"
        )

        if respuesta is True:
            self.guardar()
        elif respuesta is False:
            self.destroy()

    def guardar(self):
        persona = self.construir_persona_desde_formulario()

        if not persona.get("nombre_completo"):
            messagebox.showerror("Dato requerido", "El nombre completo es obligatorio.")
            return

        if not persona.get("nombre_corto"):
            persona["nombre_corto"] = calcular_nombre_corto(persona.get("nombre_completo"))

        if not validar_email_o_vacio(persona.get("email")):
            messagebox.showerror("E-mail inválido", "El e-mail debe tener al menos un usuario, un @ y un dominio.")
            return

        if not validar_fecha_o_vacio(persona.get("vigente_desde")):
            messagebox.showerror("Fecha inválida", "La fecha 'vigente desde' debe tener formato YYYY-MM-DD.")
            return

        if not validar_fecha_o_vacio(persona.get("vigente_hasta")):
            messagebox.showerror("Fecha inválida", "La fecha 'vigente hasta' debe tener formato YYYY-MM-DD.")
            return

        if persona.get("vigente_desde") and persona.get("vigente_hasta"):
            desde = datetime.strptime(persona["vigente_desde"], "%Y-%m-%d")
            hasta = datetime.strptime(persona["vigente_hasta"], "%Y-%m-%d")
            if desde > hasta:
                messagebox.showerror("Rango inválido", "La fecha de inicio no puede ser mayor que la fecha final.")
                return

        self.resultado = persona
        self.destroy()


# ============================================================
# Aplicación principal
# ============================================================

class ConfigurarPersonasApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Configurar Personas - Extractor Correos +")
        self.geometry("1300x720")
        self.minsize(1050, 600)

        self.data = cargar_json()
        self.personas = self.data["personas"]
        self.filtro_var = tk.StringVar()
        self.estado_var = tk.StringVar(value="Todos")
        self.hay_cambios = False
        self.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)
        self.crear_widgets()
        self.cargar_tabla()

    def crear_widgets(self):
        barra_superior = ttk.Frame(self, padding=(10, 10, 10, 4))
        barra_superior.pack(fill="x")

        ttk.Label(barra_superior, text="Buscar:").pack(side="left")
        buscar = ttk.Entry(barra_superior, textvariable=self.filtro_var, width=50)
        buscar.pack(side="left", padx=(6, 12))
        buscar.bind("<KeyRelease>", lambda _event: self.cargar_tabla())

        ttk.Label(barra_superior, text="Estado:").pack(side="left")
        estado = ttk.Combobox(barra_superior, textvariable=self.estado_var, values=["Todos", "Activos", "Inactivos", "Sin estado"], width=12, state="readonly")
        estado.pack(side="left", padx=(6, 12))
        estado.bind("<<ComboboxSelected>>", lambda _event: self.cargar_tabla())

        ttk.Button(barra_superior, text="Limpiar búsqueda", command=self.limpiar_busqueda).pack(side="left", padx=4)

        frame_tabla = ttk.Frame(self, padding=(10, 4, 10, 4))
        frame_tabla.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame_tabla, columns=COLUMNAS, show="headings", selectmode="browse")

        for col in COLUMNAS:
            self.tree.heading(col, text=COLUMNAS_VISIBLES.get(col, col), command=lambda c=col: self.ordenar_por_columna(c, False))
            ancho = 120
            if col in ["nombre_completo", "cargo", "dependencia"]:
                ancho = 240
            elif col == "email":
                ancho = 190
            elif col == "registro_id":
                ancho = 90
            elif col in ["activo", "vigente_desde", "vigente_hasta"]:
                ancho = 100
            self.tree.column(col, width=ancho, anchor="w")

        scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        frame_tabla.rowconfigure(0, weight=1)
        frame_tabla.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda _event: self.editar_persona())

        barra_botones = ttk.Frame(self, padding=(10, 4, 10, 10))
        barra_botones.pack(fill="x")

        ttk.Button(barra_botones, text="Agregar", command=self.agregar_persona).pack(side="left", padx=4)
        ttk.Button(barra_botones, text="Editar", command=self.editar_persona).pack(side="left", padx=4)
        ttk.Button(barra_botones, text="Eliminar", command=self.eliminar_persona).pack(side="left", padx=4)
        ttk.Button(barra_botones, text="Duplicar cargo", command=self.duplicar_persona).pack(side="left", padx=4)
        ttk.Button(barra_botones, text="Recargar", command=self.recargar).pack(side="left", padx=4)
        ttk.Button(barra_botones, text="Guardar", command=self.guardar).pack(side="left", padx=4)
        ttk.Button(barra_botones, text="Restaurar copia", command=self.restaurar_backup).pack(side="left", padx=4)

        ttk.Button(barra_botones, text="Cerrar", command=self.cerrar_aplicacion).pack(side="right", padx=4)

        self.label_estado = ttk.Label(self, text="", anchor="w", padding=(10, 0, 10, 8))
        self.label_estado.pack(fill="x")

    def marcar_cambios(self):
        self.hay_cambios = True
        self.title("Configurar Personas - Extractor Correos + *")

    def cerrar_aplicacion(self):
        if not self.hay_cambios:
            self.destroy()
            return

        respuesta = messagebox.askyesnocancel(
            "Cambios sin guardar",
            "Hay cambios sin guardar.\n\n¿Desea guardar antes de cerrar?"
        )

        if respuesta is True:
            guardado = self.guardar(mostrar_mensaje=False)
            if guardado:
                self.destroy()

        elif respuesta is False:
            self.destroy()

    # Si respuesta es None, cancela el cierre.

    def limpiar_busqueda(self):
        self.filtro_var.set("")
        self.estado_var.set("Todos")
        self.cargar_tabla()

    def persona_pasa_filtro(self, persona):
        filtro = normalizar_texto(self.filtro_var.get()).lower()
        estado = self.estado_var.get()

        if estado == "Activos" and persona.get("activo") is not True:
            return False
        if estado == "Inactivos" and persona.get("activo") is not False:
            return False
        if estado == "Sin estado" and persona.get("activo") is not None:
            return False

        if not filtro:
            return True

        texto = " ".join(str(persona.get(c, "")) for c in COLUMNAS).lower()
        return filtro in texto

    def cargar_tabla(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        visibles = 0
        for idx, persona in enumerate(self.personas):
            if not self.persona_pasa_filtro(persona):
                continue

            valores = []
            for col in COLUMNAS:
                if col == "activo":
                    valores.append(bool_a_texto(persona.get(col)))
                else:
                    valores.append("" if persona.get(col) is None else persona.get(col, ""))

            self.tree.insert("", "end", iid=str(idx), values=valores)
            visibles += 1

        self.label_estado.config(text=f"Mostrando {visibles} de {len(self.personas)} registros | Archivo: {PERSONAS_JSON}")

    def obtener_indice_seleccionado(self):
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Sin selección", "Seleccione un registro primero.")
            return None
        return int(seleccion[0])

    def agregar_persona(self):
        persona = {
            "registro_id": generar_id(),
            "nombre_completo": "",
            "nombre_corto": "",
            "titulo": "",
            "titulo_abreviado": "",
            "cargo": "",
            "dependencia": "",
            "email": "",
            "activo": True,
            "vigente_desde": None,
            "vigente_hasta": None,
            "area_base": "",
            "fuente": "manual",
        }

        dialog = PersonaDialog(self, "Agregar persona", persona)
        self.wait_window(dialog)

        if dialog.resultado:
            self.personas.append(dialog.resultado)
            self.marcar_cambios()
            self.cargar_tabla()

    def editar_persona(self):
        idx = self.obtener_indice_seleccionado()
        if idx is None:
            return

        dialog = PersonaDialog(self, "Editar persona", dict(self.personas[idx]))
        self.wait_window(dialog)

        if dialog.resultado:
            if dialog.resultado != self.personas[idx]:
                self.personas[idx] = dialog.resultado
                self.marcar_cambios()
                self.cargar_tabla()
            self.tree.selection_set(str(idx))

    def duplicar_persona(self):
        idx = self.obtener_indice_seleccionado()
        if idx is None:
            return

        nueva = dict(self.personas[idx])
        nueva["registro_id"] = generar_id()
        nueva["activo"] = True
        nueva["vigente_desde"] = None
        nueva["vigente_hasta"] = None
        nueva["fuente"] = "manual_duplicado"

        dialog = PersonaDialog(self, "Duplicar registro / nuevo cargo", nueva)
        self.wait_window(dialog)

        if dialog.resultado:
            self.personas.append(dialog.resultado)
            self.marcar_cambios()
            self.cargar_tabla()

    def eliminar_persona(self):
        idx = self.obtener_indice_seleccionado()
        if idx is None:
            return

        persona = self.personas[idx]
        nombre = persona.get("nombre_completo", "")
        cargo = persona.get("cargo", "")

        confirmar = messagebox.askyesno(
            "Confirmar eliminación",
            f"¿Seguro que desea eliminar este registro?\n\n{nombre}\n{cargo}\n\nEsta acción se guardará definitivamente cuando presione Guardar."
        )

        if confirmar:
            del self.personas[idx]
            self.marcar_cambios()
            self.cargar_tabla()

    def recargar(self):
        confirmar = messagebox.askyesno(
            "Recargar archivo",
            "¿Desea recargar personas.json?\n\nLos cambios no guardados se perderán."
        )
        if not confirmar:
            return

        self.data = cargar_json()
        self.personas = self.data["personas"]
        self.cargar_tabla()

    def preparar_restauracion(self):
        if not self.hay_cambios:
            return True

        respuesta = messagebox.askyesnocancel(
            "Cambios sin guardar",
            "Hay cambios sin guardar.\n\n"
            "Si continua sin guardar, esos cambios se perderan.\n\n"
            "Desea guardar los cambios antes de restaurar?"
        )

        if respuesta is None:
            return False

        if respuesta is True:
            return self.guardar(mostrar_mensaje=False)

        return True

    def restaurar_backup(self):
        if not self.preparar_restauracion():
            return

        backups = listar_backups()
        if not backups:
            messagebox.showinfo(
                "Sin copias de seguridad",
                f"No se encontraron copias de seguridad en:\n\n{BACKUP_DIR}"
            )
            return

        dialog = RestaurarCopiaDialog(self, backups)
        self.wait_window(dialog)

        if not dialog.resultado:
            return

        confirmar = messagebox.askyesno(
            "Confirmar restauracion",
            "Se restaurara esta copia de seguridad:\n\n"
            f"{dialog.resultado.name}\n\n"
            "Antes de reemplazar personas.json se creara una copia de seguridad del archivo actual.\n\n"
            "Desea continuar?"
        )
        if not confirmar:
            return

        try:
            backup_actual = crear_backup()
            shutil.copy2(dialog.resultado, PERSONAS_JSON)
            self.data = cargar_json()
            self.personas = self.data["personas"]
            self.hay_cambios = False
            self.title("Configurar Personas - Extractor Correos +")
            self.cargar_tabla()

            mensaje = "Copia de seguridad restaurada correctamente."
            if backup_actual:
                mensaje += f"\n\nCopia del estado anterior creada en:\n{backup_actual}"
            messagebox.showinfo("Restauracion completada", mensaje)

        except Exception as e:
            messagebox.showerror("Error al restaurar", f"No se pudo restaurar la copia de seguridad:\n\n{e}")

    def guardar(self, mostrar_mensaje=True):
        try:
            backup = crear_backup()
            guardar_json(self.data)

            self.hay_cambios = False
            self.title("Configurar Personas - Extractor Correos +")

            mensaje = "personas.json guardado correctamente."
            if backup:
                mensaje += f"\n\nCopia de seguridad creada:\n{backup}"

            if mostrar_mensaje:
                messagebox.showinfo("Guardado", mensaje)

            self.cargar_tabla()
            return True

        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar personas.json:\n\n{e}")
            return False

    def ordenar_por_columna(self, columna, reversa):
        def valor(idx):
            persona = self.personas[int(idx)]
            if columna == "activo":
                return bool_a_texto(persona.get(columna))
            return str(persona.get(columna, "") or "").lower()

        items = list(self.tree.get_children(""))
        items.sort(key=valor, reverse=reversa)

        for posicion, item in enumerate(items):
            self.tree.move(item, "", posicion)

        self.tree.heading(columna, command=lambda: self.ordenar_por_columna(columna, not reversa))


if __name__ == "__main__":
    app = ConfigurarPersonasApp()
    app.mainloop()
