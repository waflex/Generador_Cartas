"""Fachada del repositorio de datos.

La implementacion actual es SQLite. Para migrar a SQL Server se puede agregar
otra implementacion y seleccionarla aqui sin cambiar la interfaz.
"""

from .repositorio_sqlite import Repositorio

__all__ = ["Repositorio"]
