import json
import shutil
import uuid
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox


# ============================================================
# Configuración
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
PERSONAS_JSON = BASE_DIR / "personas.json"
BACKUP_DIR = BASE_DIR / "backups_personas"

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
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup_path = BACKUP_DIR / f"personas_backup_{timestamp}.json"
    shutil.copy2(PERSONAS_JSON, backup_path)
    return backup_path


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
        ttk.Button(botones, text="Cancelar", command=self.destroy).grid(row=0, column=1, padx=4)

        self.bind("<Return>", lambda _event: self.guardar())
        self.bind("<Escape>", lambda _event: self.destroy())

        self.update_idletasks()
        self.centrar(parent)

    def centrar(self, parent):
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def guardar(self):
        persona = dict(self.persona_original)

        for campo, var in self.variables.items():
            valor = normalizar_texto(var.get())
            persona[campo] = valor if valor else None

        if not persona.get("nombre_completo"):
            messagebox.showerror("Dato requerido", "El nombre completo es obligatorio.")
            return

        if not persona.get("nombre_corto"):
            persona["nombre_corto"] = calcular_nombre_corto(persona.get("nombre_completo"))

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

        persona["activo"] = self.activo_var.get()
        persona["registro_id"] = persona.get("registro_id") or generar_id()
        persona["fuente"] = persona.get("fuente") or "manual"

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

        ttk.Button(barra_botones, text="Cerrar", command=self.destroy).pack(side="right", padx=4)

        self.label_estado = ttk.Label(self, text="", anchor="w", padding=(10, 0, 10, 8))
        self.label_estado.pack(fill="x")

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
            self.cargar_tabla()

    def editar_persona(self):
        idx = self.obtener_indice_seleccionado()
        if idx is None:
            return

        dialog = PersonaDialog(self, "Editar persona", dict(self.personas[idx]))
        self.wait_window(dialog)

        if dialog.resultado:
            self.personas[idx] = dialog.resultado
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

    def guardar(self):
        try:
            backup = crear_backup()
            guardar_json(self.data)
            mensaje = "personas.json guardado correctamente."
            if backup:
                mensaje += f"\n\nBackup creado:\n{backup}"
            messagebox.showinfo("Guardado", mensaje)
            self.cargar_tabla()
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar personas.json:\n\n{e}")

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
