"""Conexion SQLite y utilidades de rutas."""

import sqlite3
from datetime import datetime
from pathlib import Path

import config as app_config


def ruta_bd(config=None) -> Path:
    return app_config.ruta_base_datos(config)


def conectar(config=None) -> sqlite3.Connection:
    ruta = ruta_bd(config)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(ruta))
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    con.execute("PRAGMA journal_mode = WAL")
    return con


def respaldar(config=None) -> Path:
    ruta = ruta_bd(config)
    if not ruta.exists():
        raise FileNotFoundError("La base de datos todavia no existe.")
    carpeta = app_config.carpeta_respaldos(config)
    carpeta.mkdir(parents=True, exist_ok=True)
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = carpeta / f"cruce_correos_{marca}.db"
    origen = sqlite3.connect(str(ruta))
    copia = sqlite3.connect(str(destino))
    try:
        origen.backup(copia)
    finally:
        copia.close()
        origen.close()
    return destino
