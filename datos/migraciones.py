"""Aplicacion de migraciones del esquema."""

import sqlite3
from datetime import datetime

from .esquema import ESQUEMA, VERSION


def _ahora():
    return datetime.now().isoformat(timespec="seconds")


def asegurar_esquema(con: sqlite3.Connection):
    con.execute(
        """
        CREATE TABLE IF NOT EXISTS schema_version (
            version INTEGER PRIMARY KEY,
            applied_at TEXT NOT NULL
        )
        """
    )
    fila = con.execute("SELECT MAX(version) AS version FROM schema_version").fetchone()
    actual = int(fila["version"] or 0) if isinstance(fila, sqlite3.Row) else int(fila[0] or 0)
    if actual == 0:
        con.executescript(ESQUEMA)
        con.execute(
            "INSERT OR IGNORE INTO schema_version(version, applied_at) VALUES (?, ?)",
            (VERSION, _ahora()),
        )
        con.commit()
        return VERSION
    if actual < VERSION:
        # Las migraciones futuras se agregan aqui, una por version.
        con.execute(
            "INSERT OR REPLACE INTO schema_version(version, applied_at) VALUES (?, ?)",
            (VERSION, _ahora()),
        )
        con.commit()
    return VERSION
