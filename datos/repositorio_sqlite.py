"""Repositorio SQLite usado por la interfaz y los generadores."""

from pathlib import Path

import openpyxl

import config as app_config
import generar_cartas as g

from . import db
from .fechas import parsear_rango_fechas, sugerir_periodo
from .importadores import Importador
from .migraciones import asegurar_esquema


class Repositorio:
    def __init__(self, config=None):
        self.config = config if config is not None else app_config.cargar_config()
        self.path = db.ruta_bd(self.config)
        self.con = db.conectar(self.config)
        asegurar_esquema(self.con)
        self.importador = Importador(self.con)

    def cerrar(self):
        try:
            self.con.close()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Consultas de apoyo
    # ------------------------------------------------------------------
    def _carrera_id(self, carrera):
        fila = self.con.execute(
            "SELECT id FROM carreras WHERE nombre_norm = ?",
            (g.normalizar(str(carrera or "")),),
        ).fetchone()
        return fila["id"] if fila else None

    def listar_periodos(self, carrera=None):
        if carrera:
            carrera_id = self._carrera_id(carrera)
            if not carrera_id:
                return []
            filas = self.con.execute(
                """
                SELECT DISTINCT p.nombre, p.fecha_inicio, p.fecha_fin
                FROM periodos p
                JOIN rotaciones r ON r.periodo_id = p.id
                WHERE r.carrera_id = ?
                ORDER BY p.fecha_inicio DESC, p.nombre DESC
                """,
                (carrera_id,),
            )
        else:
            filas = self.con.execute(
                """
                SELECT nombre, fecha_inicio, fecha_fin FROM periodos
                ORDER BY fecha_inicio DESC, nombre DESC
                """
            )
        return [
            {
                "nombre": fila["nombre"],
                "fecha_inicio": fila["fecha_inicio"] or "",
                "fecha_fin": fila["fecha_fin"] or "",
            }
            for fila in filas
        ]

    def _periodo(self, periodo, carrera_id):
        if periodo:
            fila = self.con.execute(
                "SELECT nombre FROM periodos WHERE nombre = ?", (periodo,)
            ).fetchone()
            return fila["nombre"] if fila else None
        fila = self.con.execute(
            """
            SELECT p.nombre FROM periodos p
            JOIN rotaciones r ON r.periodo_id = p.id
            WHERE r.carrera_id = ?
            ORDER BY p.fecha_inicio DESC, p.nombre DESC
            LIMIT 1
            """,
            (carrera_id,),
        ).fetchone()
        return fila["nombre"] if fila else None

    def listar_rotaciones(self, carrera, periodo=None):
        carrera_id = self._carrera_id(carrera)
        if not carrera_id:
            return []
        periodo = self._periodo(periodo, carrera_id)
        if not periodo:
            return []
        filas = self.con.execute(
            """
            SELECT r.codigo, r.fecha_inicio, r.fecha_fin, r.fecha_texto,
                   r.fecha_texto_normalizado, p.nombre AS periodo,
                   COUNT(a.id) AS total_alumnos
            FROM rotaciones r
            JOIN periodos p ON p.id = r.periodo_id
            LEFT JOIN asignaciones a ON a.rotacion_id = r.id
            WHERE r.carrera_id = ? AND p.nombre = ?
            GROUP BY r.id
            ORDER BY CAST(SUBSTR(r.codigo, 2) AS INTEGER), r.codigo
            """,
            (carrera_id, periodo),
        )
        return [dict(fila) for fila in filas]

    def obtener_grilla(self, carrera, periodo=None):
        carrera_id = self._carrera_id(carrera)
        if not carrera_id:
            return {}
        periodo = self._periodo(periodo, carrera_id)
        if not periodo:
            return {}
        rotaciones = self.con.execute(
            """
            SELECT r.id, r.codigo, r.fecha_inicio, r.fecha_fin, r.fecha_texto,
                   r.fecha_texto_normalizado
            FROM rotaciones r
            JOIN periodos p ON p.id = r.periodo_id
            WHERE r.carrera_id = ? AND p.nombre = ?
            ORDER BY CAST(SUBSTR(r.codigo, 2) AS INTEGER), r.codigo
            """,
            (carrera_id, periodo),
        ).fetchall()

        resultado = {}
        for rot in rotaciones:
            campos = {}
            filas = self.con.execute(
                """
                SELECT c.nombre AS campo, al.nombre AS alumno
                FROM asignaciones a
                JOIN campos_clinicos c ON c.id = a.campo_id
                JOIN alumnos al ON al.id = a.alumno_id
                WHERE a.rotacion_id = ?
                ORDER BY a.orden, a.id
                """,
                (rot["id"],),
            )
            for fila in filas:
                nombres = campos.setdefault(fila["campo"], [])
                if fila["alumno"] not in nombres:
                    nombres.append(fila["alumno"])
            resultado[rot["codigo"]] = {
                "fecha_str": rot["fecha_texto_normalizado"] or rot["fecha_texto"] or "",
                "fecha_inicio": rot["fecha_inicio"] or "",
                "fecha_fin": rot["fecha_fin"] or "",
                "campos": campos,
            }
        return resultado

    def obtener_maestro(self, carrera):
        carrera_id = self._carrera_id(carrera)
        if not carrera_id:
            return {}
        filas = self.con.execute(
            """
            SELECT nombre, nombre_norm, rut, correo, area, tipo
            FROM alumnos WHERE carrera_id = ?
            ORDER BY nombre
            """,
            (carrera_id,),
        )
        maestro = {}
        for fila in filas:
            maestro[fila["nombre_norm"]] = {
                "nombre_original": fila["nombre"],
                "rut_fmt": fila["rut"] or "",
                "correo": fila["correo"] or "",
                "area": fila["area"] or "",
                "tipo": fila["tipo"] or "",
            }
        return maestro

    def obtener_catalogo(self, carrera):
        carrera_norm = g.normalizar(str(carrera or ""))
        filas = self.con.execute(
            """
            SELECT * FROM catalogo_registros
            WHERE carrera_norm = ?
            ORDER BY campo_norm, tutor_norm, fila
            """,
            (carrera_norm,),
        )
        return [
            {
                "fila": fila["fila"],
                "tipo_practica": fila["tipo_practica"],
                "carrera": fila["carrera"],
                "area": fila["area"],
                "campo": fila["campo"],
                "campo_norm": fila["campo_norm"],
                "campo_corto": fila["campo_corto"],
                "campo_corto_norm": fila["campo_corto_norm"],
                "direccion": fila["direccion"],
                "contacto": fila["contacto"],
                "contacto_nombre": fila["contacto_nombre"],
                "tutor_texto": fila["tutor_texto"],
                "tutor": fila["tutor"],
                "tutor_norm": fila["tutor_norm"],
                "correo_tutor_texto": fila["correo_tutor_texto"],
                "correos_tutor": [
                    c.strip()
                    for c in (fila["correo_tutor"] or "").split(";")
                    if c.strip()
                ],
                "correo_tutor": fila["correo_tutor"],
                "correo_tutor_origen": fila["correo_tutor_origen"],
                "hora_ingreso": fila["hora_ingreso"],
                "nombre_residencia": fila["nombre_residencia"],
            }
            for fila in filas
        ]

    def resumen(self):
        tablas = (
            "alumnos", "rotaciones", "asignaciones", "campos_clinicos",
            "centros_clinicos", "tutores", "catalogo_registros", "importaciones",
        )
        datos = {}
        for tabla in tablas:
            fila = self.con.execute(f"SELECT COUNT(*) AS n FROM {tabla}").fetchone()
            datos[tabla] = fila["n"]
        ultima = self.con.execute(
            "SELECT * FROM importaciones ORDER BY id DESC LIMIT 1"
        ).fetchone()
        datos["ultima_importacion"] = dict(ultima) if ultima else None
        datos["archivo"] = str(self.path)
        return datos

    def historial(self, limite=50):
        filas = self.con.execute(
            """
            SELECT i.*, c.nombre AS carrera, p.nombre AS periodo
            FROM importaciones i
            LEFT JOIN carreras c ON c.id = i.carrera_id
            LEFT JOIN periodos p ON p.id = i.periodo_id
            ORDER BY i.id DESC LIMIT ?
            """,
            (limite,),
        )
        return [dict(fila) for fila in filas]

    def incidencias(self, importacion_id):
        filas = self.con.execute(
            """
            SELECT * FROM importacion_incidencias
            WHERE importacion_id = ? ORDER BY id
            """,
            (importacion_id,),
        )
        return [dict(fila) for fila in filas]

    # ------------------------------------------------------------------
    # Importacion y respaldo
    # ------------------------------------------------------------------
    def importar_alumnos(self, path, hoja, carrera):
        return self.importador.importar_alumnos(path, hoja, carrera)

    def importar_rotaciones(self, path, hoja, carrera, periodo=None):
        return self.importador.importar_rotaciones(path, hoja, carrera, periodo)

    def importar_catalogo(self, path, hoja, carrera):
        return self.importador.importar_catalogo(path, hoja, carrera)

    def sugerir_periodo_archivo(self, path, hoja=None):
        return self.importador.sugerir_periodo_archivo(path, hoja)

    def respaldar(self):
        return db.respaldar(self.config)

    def exportar_excel(self, destino):
        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        libro = openpyxl.Workbook()
        hojas = {
            "Alumnos": """
                SELECT a.id, c.nombre AS carrera, a.rut, a.nombre, a.correo,
                       a.area, a.tipo
                FROM alumnos a JOIN carreras c ON c.id = a.carrera_id
                ORDER BY c.nombre, a.nombre
            """,
            "Rotaciones": """
                SELECT r.id, p.nombre AS periodo, c.nombre AS carrera,
                       r.codigo, r.fecha_inicio, r.fecha_fin, r.fecha_texto
                FROM rotaciones r
                JOIN periodos p ON p.id = r.periodo_id
                JOIN carreras c ON c.id = r.carrera_id
                ORDER BY p.nombre, c.nombre, r.codigo
            """,
            "Asignaciones": """
                SELECT p.nombre AS periodo, c.nombre AS carrera,
                       r.codigo AS rotacion, cc.nombre AS campo, a.nombre AS alumno,
                       a.rut, a.correo
                FROM asignaciones s
                JOIN rotaciones r ON r.id = s.rotacion_id
                JOIN periodos p ON p.id = r.periodo_id
                JOIN carreras c ON c.id = r.carrera_id
                JOIN campos_clinicos cc ON cc.id = s.campo_id
                JOIN alumnos a ON a.id = s.alumno_id
                ORDER BY p.nombre, c.nombre, r.codigo, cc.nombre, a.nombre
            """,
            "Centros": """
                SELECT nombre, direccion, contacto, hora_ingreso
                FROM centros_clinicos ORDER BY nombre
            """,
            "Tutores": """
                SELECT nombre, correo, correo_origen FROM tutores ORDER BY nombre
            """,
            "Catalogo": """
                SELECT carrera, campo, direccion, contacto, tutor,
                       correo_tutor, correo_tutor_origen, hora_ingreso,
                       nombre_residencia
                FROM catalogo_registros ORDER BY carrera, campo, tutor
            """,
            "Importaciones": """
                SELECT id, tipo, archivo, hoja, fecha, filas, insertados,
                       actualizados, omitidos, errores, estado, mensaje
                FROM importaciones ORDER BY id DESC
            """,
        }
        primera = True
        for nombre, sql in hojas.items():
            ws = libro.active if primera else libro.create_sheet(nombre)
            if primera:
                ws.title = nombre
                primera = False
            filas = self.con.execute(sql)
            ws.append([columna[0] for columna in filas.description])
            for fila in filas:
                ws.append(list(fila))
            for columna in ws.columns:
                largo = max(
                    (len(str(celda.value)) for celda in columna if celda.value is not None),
                    default=10,
                )
                ws.column_dimensions[columna[0].column_letter].width = min(largo + 2, 60)
        libro.save(destino)
        return destino
