#!/usr/bin/env python3
"""
config.py
Carga y guarda la configuracion de la aplicacion.

La configuracion vive en config.json para que sea facil de editar y de
versionar. La interfaz se mantiene separada para poder migrar a SQLite mas
adelante sin cambiar el resto del sistema.
"""

import copy
import json
import os
import sys
from pathlib import Path

APP_NAME = "CruceCorreos"
CONFIG_NAME = "config.json"

DEFAULT_CONFIG = {
    "version": 1,
    "backend": "json",
    "campos_clinicos": {
        "M\u00fasculo esquel\u00e9tico CEDAM LA SERENA": "CEDAM LA SERENA",
        "M\u00fasculo esquel\u00e9tico-M\u00c1SKINE Tutor Francisco Ordenes": "M\u00c1SKINE",
        "Cardio-Respi Residencia la Araucaria": "La Araucaria",
        "Neurokin\u00e9sica Mar\u00eda \u00c1vila": "Mar\u00eda \u00c1vila",
        "Neurokinesica-SIR ROA Tutor Patricio Rodriguez": "SIR ROA",
        "CASA DE REPOSO SIR ROA": "SIR ROA",
    },
    "alias_campos": {
        "M\u00fasculo esquel\u00e9tico CEDAM LA SERENA (Tutor David Araya)": "M\u00fasculo esquel\u00e9tico CEDAM LA SERENA",
        "Cardio-Respi Residencia la Araucaria (Turor Sergio Aguirre)": "Cardio-Respi Residencia la Araucaria",
        "Neurokin\u00e9sica Mar\u00eda \u00c1vila (tutor Natanael Olcese)": "Neurokin\u00e9sica Mar\u00eda \u00c1vila",
        "Neurokinesica-SIR ROA Tutor Patricio Rodriguez (Lunes-Martes y Mi\u00e9rcoles)": "Neurokinesica-SIR ROA Tutor Patricio Rodriguez",
    },
    "outlook": {
        "modo": "grupo_presentacion_recordatorio",
        "usar_cco": False,
        "cc_tutor": True,
        "hora_ingreso": "09:00",
    },
    "recordatorio": {
        "habilitado": True,
        "hora_ingreso": "09:00",
    },
    "correo_tutor_overrides": {},
}


def _es_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_dir() -> Path:
    """Carpeta escribible de la aplicacion (junto al exe o al script)."""
    if _es_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_dir() -> Path:
    """Carpeta de recursos empaquetados por PyInstaller."""
    if _es_frozen():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
    return Path(__file__).resolve().parent


def appdata_dir() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / APP_NAME
    return Path.home() / f".{APP_NAME.lower()}"


def config_candidates() -> list:
    candidatos = []
    env = os.environ.get("CRUCE_CORREOS_CONFIG")
    if env:
        candidatos.append(Path(env))
    candidatos.append(app_dir() / CONFIG_NAME)
    candidatos.append(appdata_dir() / CONFIG_NAME)
    candidatos.append(resource_dir() / CONFIG_NAME)
    vistos = []
    for ruta in candidatos:
        if ruta not in vistos:
            vistos.append(ruta)
    return vistos


def config_path() -> Path:
    """Devuelve la primera configuracion existente o la ruta preferida."""
    for ruta in config_candidates():
        if ruta.exists():
            return ruta
    return app_dir() / CONFIG_NAME


def _leer_json(ruta: Path) -> dict:
    try:
        with ruta.open("r", encoding="utf-8") as archivo:
            data = json.load(archivo)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _mezclar(base: dict, datos: dict) -> dict:
    resultado = copy.deepcopy(base)
    for clave, valor in datos.items():
        if clave in ("outlook", "recordatorio") and isinstance(valor, dict):
            actual = resultado.get(clave, {})
            if isinstance(actual, dict):
                actual.update(valor)
                resultado[clave] = actual
            else:
                resultado[clave] = valor
        else:
            resultado[clave] = valor
    return resultado


def cargar_config() -> dict:
    for ruta in config_candidates():
        if ruta.exists():
            return _mezclar(DEFAULT_CONFIG, _leer_json(ruta))
    return copy.deepcopy(DEFAULT_CONFIG)


def _ruta_escribible() -> Path:
    env = os.environ.get("CRUCE_CORREOS_CONFIG")
    if env:
        ruta = Path(env)
        ruta.parent.mkdir(parents=True, exist_ok=True)
        return ruta

    preferida = app_dir() / CONFIG_NAME
    try:
        preferida.parent.mkdir(parents=True, exist_ok=True)
        if not preferida.exists():
            preferida.touch()
        return preferida
    except OSError:
        pass

    alternativa = appdata_dir() / CONFIG_NAME
    alternativa.parent.mkdir(parents=True, exist_ok=True)
    return alternativa


def guardar_config(config: dict) -> Path:
    ruta = _ruta_escribible()
    with ruta.open("w", encoding="utf-8") as archivo:
        json.dump(config, archivo, ensure_ascii=False, indent=2)
        archivo.write("\n")
    return ruta


def asegurar_config() -> Path:
    """Crea config.json con los valores por defecto si aun no existe."""
    for ruta in config_candidates():
        if ruta.exists():
            return ruta
    return guardar_config(copy.deepcopy(DEFAULT_CONFIG))


def ruta_recurso(nombre: str) -> Path:
    """Busca un recurso junto al exe/script o dentro del paquete."""
    for base in (app_dir(), resource_dir()):
        ruta = base / nombre
        if ruta.exists():
            return ruta
    return app_dir() / nombre
