"""Importadores de los Excel del sistema hacia SQLite."""

import hashlib
import json
from datetime import datetime
from pathlib import Path

import catalogo_tutores as ct
import generar_cartas as g

from .fechas import parsear_rango_fechas, sugerir_periodo


def _ahora():
    return datetime.now().isoformat(timespec="seconds")


def _texto(valor):
    return "" if valor is None else str(valor).strip()


def _norm(valor):
    return g.normalizar(_texto(valor))


def _norm_rut(valor):
    return "".join(c for c in _texto(valor).upper() if c.isdigit() or c == "K")


def _hash_archivo(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as archivo:
        for bloque in iter(lambda: archivo.read(1024 * 1024), b""):
            h.update(bloque)
    return h.hexdigest()


class Importador:
    def __init__(self, con):
        self.con = con

    # ------------------------------------------------------------------
    # Catalogos base
    # ------------------------------------------------------------------
    def _carrera(self, nombre):
        nombre = _texto(nombre) or "SIN CARRERA"
        nombre_norm = _norm(nombre)
        fila = self.con.execute(
            "SELECT id FROM carreras WHERE nombre_norm = ?", (nombre_norm,)
        ).fetchone()
        if fila:
            return fila["id"]
        cur = self.con.execute(
            "INSERT INTO carreras(nombre, nombre_norm) VALUES (?, ?)",
            (nombre, nombre_norm),
        )
        return cur.lastrowid

    def _periodo(self, nombre, fecha_inicio="", fecha_fin=""):
        nombre = _texto(nombre) or sugerir_periodo(fecha_inicio, fecha_fin)
        fila = self.con.execute(
            "SELECT id, fecha_inicio, fecha_fin FROM periodos WHERE nombre = ?",
            (nombre,),
        ).fetchone()
        if fila:
            inicios = [v for v in (fila["fecha_inicio"], fecha_inicio) if v]
            fines = [v for v in (fila["fecha_fin"], fecha_fin) if v]
            inicio = min(inicios) if inicios else ""
            fin = max(fines) if fines else ""
            self.con.execute(
                "UPDATE periodos SET fecha_inicio = ?, fecha_fin = ? WHERE id = ?",
                (inicio or "", fin or "", fila["id"]),
            )
            return fila["id"]
        cur = self.con.execute(
            "INSERT INTO periodos(nombre, fecha_inicio, fecha_fin) VALUES (?, ?, ?)",
            (nombre, fecha_inicio or "", fecha_fin or ""),
        )
        return cur.lastrowid

    def _campo(self, carrera_id, nombre):
        nombre = _texto(nombre)
        nombre_norm = _norm(nombre)
        fila = self.con.execute(
            "SELECT id FROM campos_clinicos WHERE carrera_id = ? AND nombre_norm = ?",
            (carrera_id, nombre_norm),
        ).fetchone()
        if fila:
            self.con.execute(
                "UPDATE campos_clinicos SET nombre = ?, nombre_corto = ? WHERE id = ?",
                (nombre, g.resumir_campo(nombre), fila["id"]),
            )
            return fila["id"], False
        cur = self.con.execute(
            """
            INSERT INTO campos_clinicos(carrera_id, nombre, nombre_norm, nombre_corto)
            VALUES (?, ?, ?, ?)
            """,
            (carrera_id, nombre, nombre_norm, g.resumir_campo(nombre)),
        )
        return cur.lastrowid, True

    def _centro(self, nombre, direccion="", contacto="", hora=""):
        nombre = _texto(nombre)
        nombre_norm = _norm(nombre)
        direccion = _texto(direccion)
        direccion_norm = _norm(direccion)
        fila = self.con.execute(
            """
            SELECT id FROM centros_clinicos
            WHERE nombre_norm = ? AND direccion_norm = ?
            """,
            (nombre_norm, direccion_norm),
        ).fetchone()
        ahora = _ahora()
        if fila:
            self.con.execute(
                """
                UPDATE centros_clinicos
                SET nombre = ?, direccion = ?, contacto = ?, hora_ingreso = ?,
                    actualizado_en = ?
                WHERE id = ?
                """,
                (nombre, direccion, _texto(contacto), _texto(hora), ahora, fila["id"]),
            )
            return fila["id"], False
        cur = self.con.execute(
            """
            INSERT INTO centros_clinicos(
                nombre, nombre_norm, direccion, direccion_norm, contacto,
                hora_ingreso, creado_en, actualizado_en
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                nombre, nombre_norm, direccion, direccion_norm,
                _texto(contacto), _texto(hora), ahora, ahora,
            ),
        )
        return cur.lastrowid, True

    def _tutor(self, nombre, correo="", origen=""):
        nombre = _texto(nombre)
        nombre_norm = _norm(nombre)
        correo = _texto(correo)
        fila = self.con.execute(
            "SELECT id FROM tutores WHERE nombre_norm = ? AND correo = ?",
            (nombre_norm, correo),
        ).fetchone()
        ahora = _ahora()
        if fila:
            self.con.execute(
                """
                UPDATE tutores SET nombre = ?, correo_origen = ?, actualizado_en = ?
                WHERE id = ?
                """,
                (nombre, _texto(origen), ahora, fila["id"]),
            )
            return fila["id"], False
        cur = self.con.execute(
            """
            INSERT INTO tutores(
                nombre, nombre_norm, correo, correo_origen, creado_en, actualizado_en
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (nombre, nombre_norm, correo, _texto(origen), ahora, ahora),
        )
        return cur.lastrowid, True

    def _alumno_por_nombre(self, carrera_id, nombre, mapa=None):
        nombre_norm = _norm(nombre)
        if mapa is not None:
            alumno_id = mapa.get(nombre_norm)
            if alumno_id:
                return alumno_id, "exacto"
            encontrado, tipo = g.buscar_alumno(nombre, mapa)
            if encontrado:
                return encontrado, tipo
            return None, None
        fila = self.con.execute(
            "SELECT id FROM alumnos WHERE carrera_id = ? AND nombre_norm = ?",
            (carrera_id, nombre_norm),
        ).fetchone()
        return (fila["id"], "exacto") if fila else (None, None)

    def _upsert_alumno(self, carrera_id, nombre, rut="", correo="", area="", tipo=""):
        nombre = _texto(nombre)
        nombre_norm = _norm(nombre)
        rut = _texto(rut)
        rut_norm = _norm_rut(rut)
        clave = rut_norm or f"nombre:{nombre_norm}"
        ahora = _ahora()

        fila = self.con.execute(
            "SELECT id FROM alumnos WHERE carrera_id = ? AND clave = ?",
            (carrera_id, clave),
        ).fetchone()

        if not fila and rut_norm:
            fila = self.con.execute(
                """
                SELECT id FROM alumnos
                WHERE carrera_id = ? AND nombre_norm = ? AND clave LIKE 'nombre:%'
                ORDER BY id LIMIT 1
                """,
                (carrera_id, nombre_norm),
            ).fetchone()
            if fila:
                ocupado = self.con.execute(
                    "SELECT id FROM alumnos WHERE carrera_id = ? AND clave = ?",
                    (carrera_id, clave),
                ).fetchone()
                if ocupado:
                    self.con.execute(
                        "UPDATE OR IGNORE asignaciones SET alumno_id = ? WHERE alumno_id = ?",
                        (ocupado["id"], fila["id"]),
                    )
                    self.con.execute(
                        "DELETE FROM asignaciones WHERE alumno_id = ?", (fila["id"],)
                    )
                    self.con.execute("DELETE FROM alumnos WHERE id = ?", (fila["id"],))
                    fila = ocupado
                else:
                    self.con.execute(
                        """
                        UPDATE alumnos
                        SET rut = ?, rut_norm = ?, clave = ?, nombre = ?,
                            nombre_norm = ?, correo = ?, area = ?, tipo = ?,
                            actualizado_en = ?
                        WHERE id = ?
                        """,
                        (
                            rut, rut_norm, clave, nombre, nombre_norm, _texto(correo),
                            _texto(area), _texto(tipo), ahora, fila["id"],
                        ),
                    )
                    return fila["id"], False

        if fila:
            self.con.execute(
                """
                UPDATE alumnos
                SET rut = ?, rut_norm = ?, nombre = ?, nombre_norm = ?,
                    correo = ?, area = ?, tipo = ?, actualizado_en = ?
                WHERE id = ?
                """,
                (
                    rut, rut_norm, nombre, nombre_norm, _texto(correo),
                    _texto(area), _texto(tipo), ahora, fila["id"],
                ),
            )
            return fila["id"], False

        cur = self.con.execute(
            """
            INSERT INTO alumnos(
                carrera_id, rut, rut_norm, nombre, nombre_norm, correo, area,
                tipo, clave, creado_en, actualizado_en
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                carrera_id, rut, rut_norm, nombre, nombre_norm, _texto(correo),
                _texto(area), _texto(tipo), clave, ahora, ahora,
            ),
        )
        return cur.lastrowid, True

    # ------------------------------------------------------------------
    # Registro de importaciones
    # ------------------------------------------------------------------
    def _iniciar(self, tipo, path, hoja, carrera_id, periodo_id=None):
        cur = self.con.execute(
            """
            INSERT INTO importaciones(
                tipo, archivo, hoja, hash_archivo, periodo_id, carrera_id,
                fecha, estado
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'en_proceso')
            """,
            (
                tipo, str(Path(path)), _texto(hoja), _hash_archivo(path),
                periodo_id, carrera_id, _ahora(),
            ),
        )
        self.con.commit()
        return cur.lastrowid

    def _finalizar(self, importacion_id, filas, insertados, actualizados,
                   omitidos, errores, mensaje=""):
        estado = "error" if errores and not insertados and not actualizados else "ok"
        if errores:
            estado = "con_incidencias" if estado == "ok" else "error"
        self.con.execute(
            """
            UPDATE importaciones
            SET filas = ?, insertados = ?, actualizados = ?, omitidos = ?,
                errores = ?, estado = ?, mensaje = ?
            WHERE id = ?
            """,
            (
                filas, insertados, actualizados, omitidos, errores, estado,
                _texto(mensaje), importacion_id,
            ),
        )
        self.con.commit()
        return estado

    def _incidencia(self, importacion_id, fila, columna, valor, mensaje,
                    severidad="advertencia"):
        self.con.execute(
            """
            INSERT INTO importacion_incidencias(
                importacion_id, numero_fila, columna, valor, severidad, mensaje
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (importacion_id, fila, _texto(columna), _texto(valor), severidad, mensaje),
        )

    # ------------------------------------------------------------------
    # Importaciones
    # ------------------------------------------------------------------
    def importar_alumnos(self, path, hoja, carrera):
        path = Path(path)
        carrera_id = self._carrera(carrera)
        imp_id = self._iniciar("alumnos", path, hoja, carrera_id)
        insertados = actualizados = 0
        filas = 0
        try:
            maestro = g.cargar_maestro(path, hoja)
            with self.con:
                for alumno in maestro.values():
                    _, nuevo = self._upsert_alumno(
                        carrera_id,
                        alumno.get("nombre_original", ""),
                        alumno.get("rut_fmt", ""),
                        alumno.get("correo", ""),
                        alumno.get("area", ""),
                        alumno.get("tipo", ""),
                    )
                    filas += 1
                    if nuevo:
                        insertados += 1
                    else:
                        actualizados += 1
            estado = self._finalizar(imp_id, filas, insertados, actualizados, 0, 0)
            return {
                "importacion_id": imp_id,
                "tipo": "alumnos",
                "filas": filas,
                "insertados": insertados,
                "actualizados": actualizados,
                "omitidos": 0,
                "errores": 0,
                "estado": estado,
            }
        except Exception as exc:
            self._finalizar(imp_id, filas, insertados, actualizados, 0, 1, str(exc))
            raise

    def sugerir_periodo_archivo(self, path, hoja=None):
        grilla = g.parsear_grilla(Path(path), hoja)
        for info in grilla.values():
            if info.get("fecha_inicio") or info.get("fecha_str"):
                return sugerir_periodo(
                    info.get("fecha_inicio") or info.get("fecha_str"),
                    info.get("fecha_fin"),
                )
        return sugerir_periodo("", "")

    def importar_rotaciones(self, path, hoja, carrera, periodo=None):
        path = Path(path)
        carrera_id = self._carrera(carrera)
        grilla = g.parsear_grilla(path, hoja)
        if not grilla:
            raise ValueError("No se encontraron rotaciones en la planilla.")

        fechas = {}
        for codigo, info in grilla.items():
            rango = parsear_rango_fechas(
                info.get("fecha_texto") or info.get("fecha_str") or ""
            )
            if info.get("fecha_inicio"):
                rango["fecha_inicio"] = info["fecha_inicio"]
            if info.get("fecha_fin"):
                rango["fecha_fin"] = info["fecha_fin"]
            fechas[codigo] = rango

        primero = next(iter(fechas.values()), {})
        periodo = periodo or sugerir_periodo(
            primero.get("fecha_inicio"), primero.get("fecha_fin")
        )
        periodo_id = self._periodo(
            periodo, primero.get("fecha_inicio", ""), primero.get("fecha_fin", "")
        )
        imp_id = self._iniciar(
            "rotaciones", path, hoja, carrera_id, periodo_id
        )
        insertados = actualizados = omitidos = errores = 0
        filas = 0
        try:
            with self.con:
                mapa = {
                    fila["nombre_norm"]: fila["id"]
                    for fila in self.con.execute(
                        "SELECT id, nombre_norm FROM alumnos WHERE carrera_id = ?",
                        (carrera_id,),
                    )
                }
                for codigo in g.ordenar_rotaciones(grilla):
                    info = grilla[codigo]
                    rango = fechas[codigo]
                    filas += 1
                    if not rango["fecha_inicio"]:
                        errores += 1
                        self._incidencia(
                            imp_id, None, "fecha_inicio", info.get("fecha_str", ""),
                            "Rotacion sin fecha de inicio.",
                        )
                    if not rango["fecha_fin"]:
                        self._incidencia(
                            imp_id, None, "fecha_fin", info.get("fecha_str", ""),
                            "Rotacion sin fecha de termino.",
                        )
                    if (
                        rango["fecha_inicio"] and rango["fecha_fin"]
                        and rango["fecha_inicio"] > rango["fecha_fin"]
                    ):
                        errores += 1
                        self._incidencia(
                            imp_id, None, "fecha_fin", info.get("fecha_str", ""),
                            "La fecha de termino es anterior a la de inicio.",
                            severidad="error",
                        )

                    fila_rot = self.con.execute(
                        """
                        SELECT id FROM rotaciones
                        WHERE periodo_id = ? AND carrera_id = ? AND codigo = ?
                        """,
                        (periodo_id, carrera_id, codigo),
                    ).fetchone()
                    ahora = _ahora()
                    datos = (
                        rango["fecha_inicio"], rango["fecha_fin"],
                        rango["fecha_texto"],
                        rango["fecha_texto_normalizado"], ahora,
                    )
                    if fila_rot:
                        rotacion_id = fila_rot["id"]
                        self.con.execute(
                            """
                            UPDATE rotaciones
                            SET fecha_inicio = ?, fecha_fin = ?, fecha_texto = ?,
                                fecha_texto_normalizado = ?, actualizado_en = ?
                            WHERE id = ?
                            """,
                            datos + (rotacion_id,),
                        )
                        actualizados += 1
                    else:
                        cur = self.con.execute(
                            """
                            INSERT INTO rotaciones(
                                periodo_id, carrera_id, codigo, fecha_inicio,
                                fecha_fin, fecha_texto, fecha_texto_normalizado,
                                creado_en, actualizado_en
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                periodo_id, carrera_id, codigo,
                                rango["fecha_inicio"], rango["fecha_fin"],
                                rango["fecha_texto"],
                                rango["fecha_texto_normalizado"], ahora, ahora,
                            ),
                        )
                        rotacion_id = cur.lastrowid
                        insertados += 1

                    self.con.execute(
                        "DELETE FROM asignaciones WHERE rotacion_id = ?",
                        (rotacion_id,),
                    )
                    for campo_orden, (campo, nombres) in enumerate(
                        info.get("campos", {}).items()
                    ):
                        campo_id, _ = self._campo(carrera_id, campo)
                        for alumno_orden, nombre in enumerate(nombres):
                            orden = campo_orden * 1000 + alumno_orden
                            alumno_id, _ = self._alumno_por_nombre(
                                carrera_id, nombre, mapa
                            )
                            if not alumno_id:
                                alumno_id, _ = self._upsert_alumno(
                                    carrera_id, nombre
                                )
                                mapa[_norm(nombre)] = alumno_id
                                self._incidencia(
                                    imp_id, None, "alumno", nombre,
                                    "Alumno sin registro previo; se creo un "
                                    "registro pendiente de completar.",
                                )
                            cur = self.con.execute(
                                """
                                INSERT OR IGNORE INTO asignaciones(
                                    rotacion_id, alumno_id, campo_id, orden
                                ) VALUES (?, ?, ?, ?)
                                """,
                                (rotacion_id, alumno_id, campo_id, orden),
                            )
                            if cur.rowcount == 0:
                                omitidos += 1
                inicios = [
                    r["fecha_inicio"] for r in fechas.values() if r["fecha_inicio"]
                ]
                fines = [r["fecha_fin"] for r in fechas.values() if r["fecha_fin"]]
                self.con.execute(
                    """
                    UPDATE periodos SET fecha_inicio = ?, fecha_fin = ?
                    WHERE id = ?
                    """,
                    (
                        min(inicios) if inicios else "",
                        max(fines) if fines else "",
                        periodo_id,
                    ),
                )
            estado = self._finalizar(
                imp_id, filas, insertados, actualizados, omitidos, errores
            )
            return {
                "importacion_id": imp_id,
                "tipo": "rotaciones",
                "periodo": periodo,
                "filas": filas,
                "insertados": insertados,
                "actualizados": actualizados,
                "omitidos": omitidos,
                "errores": errores,
                "estado": estado,
            }
        except Exception as exc:
            self._finalizar(
                imp_id, filas, insertados, actualizados, omitidos, errores or 1,
                str(exc),
            )
            raise

    def importar_catalogo(self, path, hoja, carrera):
        path = Path(path)
        carrera_id = self._carrera(carrera)
        registros = ct.leer_catalogo(path, hoja)
        if not registros:
            raise ValueError("No se encontraron tutores en el catalogo.")

        imp_id = self._iniciar("catalogo", path, hoja, carrera_id)
        insertados = actualizados = 0
        filas = 0
        try:
            carrera_norm = _norm(carrera)
            with self.con:
                self.con.execute(
                    "DELETE FROM catalogo_registros WHERE carrera_norm = ?",
                    (carrera_norm,),
                )
                for reg in registros:
                    filas += 1
                    campo = reg.get("campo", "")
                    campo_id, nuevo_campo = self._campo(carrera_id, campo)
                    centro_id, nuevo_centro = self._centro(
                        reg.get("nombre_residencia") or reg.get("campo_corto") or campo,
                        reg.get("direccion", ""),
                        reg.get("contacto", ""),
                        reg.get("hora_ingreso", ""),
                    )
                    tutor_id = None
                    nuevo_tutor = False
                    if reg.get("tutor"):
                        tutor_id, nuevo_tutor = self._tutor(
                            reg.get("tutor", ""),
                            reg.get("correo_tutor", ""),
                            reg.get("correo_tutor_origen", ""),
                        )
                        self.con.execute(
                            """
                            INSERT OR IGNORE INTO centro_tutores(
                                centro_id, tutor_id, campo_id, principal
                            ) VALUES (?, ?, ?, 1)
                            """,
                            (centro_id, tutor_id, campo_id),
                        )
                    self.con.execute(
                        """
                        INSERT INTO catalogo_registros(
                            importacion_id, carrera, carrera_norm, fila,
                            tipo_practica, area, campo, campo_norm, campo_corto,
                            campo_corto_norm, direccion, contacto,
                            contacto_nombre, tutor_texto, tutor, tutor_norm,
                            correo_tutor_texto, correo_tutor,
                            correo_tutor_origen, hora_ingreso,
                            nombre_residencia, centro_id, tutor_id, campo_id
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                  ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            imp_id, carrera, carrera_norm, reg.get("fila"),
                            reg.get("tipo_practica", ""), reg.get("area", ""),
                            campo, reg.get("campo_norm", ""),
                            reg.get("campo_corto", ""),
                            reg.get("campo_corto_norm", ""),
                            reg.get("direccion", ""), reg.get("contacto", ""),
                            reg.get("contacto_nombre", ""),
                            reg.get("tutor_texto", ""), reg.get("tutor", ""),
                            reg.get("tutor_norm", ""),
                            reg.get("correo_tutor_texto", ""),
                            reg.get("correo_tutor", ""),
                            reg.get("correo_tutor_origen", ""),
                            reg.get("hora_ingreso", ""),
                            reg.get("nombre_residencia", ""),
                            centro_id, tutor_id, campo_id,
                        ),
                    )
                    if nuevo_campo or nuevo_centro or nuevo_tutor:
                        insertados += 1
                    else:
                        actualizados += 1
            estado = self._finalizar(imp_id, filas, insertados, actualizados, 0, 0)
            return {
                "importacion_id": imp_id,
                "tipo": "catalogo",
                "filas": filas,
                "insertados": insertados,
                "actualizados": actualizados,
                "omitidos": 0,
                "errores": 0,
                "estado": estado,
            }
        except Exception as exc:
            self._finalizar(imp_id, filas, insertados, actualizados, 0, 1, str(exc))
            raise
