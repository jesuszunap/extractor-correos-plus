import json
import os
import shutil
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
CONFIG_PATH = BASE_DIR / "config_storage.json"
LOCAL_PERSONAS_JSON = BASE_DIR / "personas.json"
ONEDRIVE_FOLDER_NAME = "Servidores DIOU"


def cargar_config_storage():
    if not CONFIG_PATH.exists():
        return {"personas_storage": "local"}

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {"personas_storage": "local"}

    if not isinstance(data, dict):
        return {"personas_storage": "local"}

    storage = str(data.get("personas_storage", "local")).strip().lower()
    if storage not in {"local", "onedrive"}:
        storage = "local"

    return {"personas_storage": storage}


def guardar_config_storage(storage):
    storage = str(storage or "local").strip().lower()
    if storage not in {"local", "onedrive"}:
        storage = "local"

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump({"personas_storage": storage}, f, ensure_ascii=False, indent=2)


def obtener_carpetas_onedrive():
    candidatos = []

    for variable in ("OneDriveCommercial", "OneDriveConsumer", "OneDrive"):
        valor = os.environ.get(variable)
        if valor:
            candidatos.append(Path(valor))

    perfil = os.environ.get("USERPROFILE")
    if perfil:
        perfil_path = Path(perfil)
        candidatos.extend(perfil_path.glob("OneDrive*"))

    unicos = []
    vistos = set()
    for candidato in candidatos:
        try:
            resuelto = candidato.expanduser().resolve()
        except Exception:
            continue

        if resuelto in vistos:
            continue

        vistos.add(resuelto)
        if resuelto.exists() and resuelto.is_dir():
            unicos.append(resuelto)

    return unicos


def obtener_documentos_onedrive():
    for onedrive in obtener_carpetas_onedrive():
        for nombre in ("Documents", "Documentos"):
            documentos = onedrive / nombre
            if documentos.exists() and documentos.is_dir():
                return documentos

    carpetas = obtener_carpetas_onedrive()
    if carpetas:
        return carpetas[0] / "Documents"

    return None


def obtener_personas_json_onedrive():
    documentos = obtener_documentos_onedrive()
    if documentos is None:
        return None

    return documentos / ONEDRIVE_FOLDER_NAME / "personas.json"


def preparar_personas_onedrive():
    personas_onedrive = obtener_personas_json_onedrive()
    if personas_onedrive is None:
        return None

    try:
        personas_onedrive.parent.mkdir(parents=True, exist_ok=True)

        if not personas_onedrive.exists() and LOCAL_PERSONAS_JSON.exists():
            shutil.copy2(LOCAL_PERSONAS_JSON, personas_onedrive)

        return personas_onedrive
    except Exception:
        return None


def obtener_personas_json():
    config = cargar_config_storage()

    if config["personas_storage"] == "onedrive":
        personas_onedrive = preparar_personas_onedrive()
        if personas_onedrive is not None:
            return personas_onedrive

    return LOCAL_PERSONAS_JSON


def describir_ubicacion_personas():
    path = obtener_personas_json()
    modo = "OneDrive" if path != LOCAL_PERSONAS_JSON else "local"
    return modo, path
