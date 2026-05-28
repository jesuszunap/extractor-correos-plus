from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
LOCAL_PERSONAS_JSON = BASE_DIR / "personas.json"


def obtener_personas_json():
    return LOCAL_PERSONAS_JSON


def describir_ubicacion_personas():
    path = obtener_personas_json()
    return "local", path
