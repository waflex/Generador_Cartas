#!/usr/bin/env python3
"""
catalogo_tutores.py
Lee el tercer Excel de campos clinicos, direcciones, contactos y tutores.

Columnas esperadas:
    Tipo Practica | Carrera | Area | Campo Clinico | Direccion | Contacto | Tutor(a)

Columnas opcionales:
    Correo_Tutor | Hora_Ingreso | Nombre_Residencia

El correo del tutor se toma, en este orden:
    1. Columna Correo_Tutor (si tiene un correo valido).
    2. Correo incluido en Tutor(a).
    3. Correo incluido en Contacto.
"""

import re
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl

import generar_cartas as g

RE_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def _texto(valor) -> str:
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def _norm(valor) -> str:
    return g.normalizar(_texto(valor))


def extraer_correos(valor) -> list:
    """Extrae correos unicos de un texto."""
    correos = []
    for correo in RE_EMAIL.findall(str(valor or "")):
        correo = correo.strip().strip(".,;:")
        if correo and correo.lower() not in [c.lower() for c in correos]:
            correos.append(correo)
    return correos


def limpiar_nombre(valor) -> str:
    """Quita correos y etiquetas para dejar el nombre visible."""
    texto = RE_EMAIL.sub(" ", str(valor or ""))
    texto = re.sub(
        r"\b(correo|email|mail|contacto|tutor[a]?)\s*:?\s*", " ", texto,
        flags=re.IGNORECASE,
    )
    texto = re.sub(r"[;|/<>()\[\]{}]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip(" .,-")


def _columnas(encabezados) -> dict:
    cols = {}
    for indice, valor in enumerate(encabezados):
        normal = _norm(valor).replace("(A)", " ").replace("(A)", " ")
        normal = re.sub(r"\s+", " ", normal).strip()
        if not normal:
            continue

        if "CORREO" in normal or "EMAIL" in normal or "MAIL" in normal:
            if "TUTOR" in normal:
                cols.setdefault("correo_tutor", indice)
                continue
        if "TIPO" in normal and "PRACTICA" in normal:
            cols.setdefault("tipo", indice)
        elif normal == "TIPO" or normal.startswith("TIPO "):
            cols.setdefault("tipo", indice)
        if "CARRERA" in normal:
            cols.setdefault("carrera", indice)
        if normal == "AREA" or normal.startswith("AREA "):
            cols.setdefault("area", indice)
        if "CAMPO" in normal and "CLINICO" in normal:
            cols.setdefault("campo", indice)
        if "DIRECCION" in normal:
            cols.setdefault("direccion", indice)
        if "CONTACTO" in normal:
            cols.setdefault("contacto", indice)
        if "TUTOR" in normal and "CORREO" not in normal and "EMAIL" not in normal:
            cols.setdefault("tutor", indice)
        if "HORA" in normal and ("INGRESO" in normal or "ENTRADA" in normal):
            cols.setdefault("hora", indice)
        if "RESIDENCIA" in normal and ("NOMBRE" in normal or "CENTRO" in normal):
            cols.setdefault("nombre_residencia", indice)
        if normal in ("NOMBRE RESIDENCIA", "NOMBRE DEL CENTRO", "CENTRO"):
            cols.setdefault("nombre_residencia", indice)
    return cols


def detectar_columnas(ws):
    """Devuelve (columnas, fila_encabezado)."""
    filas = list(ws.iter_rows(values_only=True))
    for indice, fila in enumerate(filas[:40]):
        cols = _columnas(fila)
        if "campo" in cols and ("tutor" in cols or "contacto" in cols):
            return cols, indice + 1
    for indice, fila in enumerate(filas):
        cols = _columnas(fila)
        if "campo" in cols and ("tutor" in cols or "contacto" in cols):
            return cols, indice + 1
    raise ValueError(
        "No se encontro el encabezado del catalogo. Se esperaban columnas "
        "como Campo Clinico, Direccion, Contacto y Tutor(a)."
    )


def _valor(fila, cols, clave):
    indice = cols.get(clave)
    if indice is None or indice >= len(fila):
        return ""
    return _texto(fila[indice])


def leer_catalogo(path: Path, hoja=None) -> list:
    """Lee el catalogo de tutores y devuelve una lista de registros."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = g.elegir_hoja(wb, hoja)
        cols, primera = detectar_columnas(ws)
        registros = []
        for numero, fila in enumerate(
            ws.iter_rows(min_row=primera + 1, values_only=True), start=primera + 1
        ):
            campo = _valor(fila, cols, "campo")
            if not campo:
                continue

            tutor_texto = _valor(fila, cols, "tutor")
            contacto = _valor(fila, cols, "contacto")
            correo_columna = _valor(fila, cols, "correo_tutor")

            tutor = limpiar_nombre(tutor_texto)
            contacto_nombre = limpiar_nombre(contacto)

            correos = extraer_correos(correo_columna)
            origen = "columna" if correos else ""
            if not correos:
                correos_tutor = extraer_correos(tutor_texto)
                correos_contacto = extraer_correos(contacto)
                correos = []
                for correo in correos_tutor + correos_contacto:
                    if correo.lower() not in [c.lower() for c in correos]:
                        correos.append(correo)
                if correos_tutor and correos_contacto:
                    origen = "tutor+contacto"
                elif correos_tutor:
                    origen = "tutor"
                elif correos_contacto:
                    origen = "contacto"

            if not tutor:
                tutor = contacto_nombre

            nombre_residencia = _valor(fila, cols, "nombre_residencia")
            if not nombre_residencia:
                nombre_residencia = g.resumir_campo(campo)

            registros.append({
                "fila": numero,
                "tipo_practica": _valor(fila, cols, "tipo"),
                "carrera": _valor(fila, cols, "carrera"),
                "area": _valor(fila, cols, "area"),
                "campo": campo,
                "campo_norm": _norm(g.limpiar_campo(campo)),
                "campo_corto": g.resumir_campo(campo),
                "campo_corto_norm": _norm(g.resumir_campo(campo)),
                "direccion": _valor(fila, cols, "direccion"),
                "contacto": contacto,
                "contacto_nombre": contacto_nombre,
                "tutor_texto": tutor_texto,
                "tutor": tutor,
                "tutor_norm": _norm(tutor),
                "correo_tutor_texto": correo_columna,
                "correos_tutor": correos,
                "correo_tutor": "; ".join(correos),
                "correo_tutor_origen": origen,
                "hora_ingreso": _valor(fila, cols, "hora"),
                "nombre_residencia": nombre_residencia,
            })
        return registros
    finally:
        wb.close()


def _puntaje(registro, campo_norm, corto_norm, tutor_norm, carrera, tipo):
    puntaje = 0
    motivo = ""

    campo_reg = registro.get("campo_norm", "")
    corto_reg = registro.get("campo_corto_norm", "")
    tutor_reg = registro.get("tutor_norm", "")

    if campo_norm and campo_reg == campo_norm:
        puntaje, motivo = 100, "campo exacto"
    elif corto_norm and corto_reg == corto_norm:
        puntaje, motivo = 95, "nombre corto"
    elif campo_norm and campo_reg and (
        campo_reg in campo_norm or campo_norm in campo_reg
    ):
        puntaje, motivo = 82, "campo parcial"

    if tutor_norm and tutor_reg:
        if tutor_norm == tutor_reg:
            puntaje, motivo = max(puntaje, 90), "tutor exacto"
        elif tutor_norm in tutor_reg or tutor_reg in tutor_norm:
            puntaje, motivo = max(puntaje, 76), "tutor parcial"

    for valor_a, valor_b, base in (
        (campo_norm, campo_reg, 84),
        (corto_norm, corto_reg, 88),
        (tutor_norm, tutor_reg, 86),
    ):
        if not valor_a or not valor_b:
            continue
        ratio = SequenceMatcher(None, valor_a, valor_b).ratio()
        if ratio >= 0.86:
            puntaje = max(puntaje, round(base * ratio))
            if not motivo:
                motivo = "similitud"

    if not puntaje:
        return 0, ""

    carrera_reg = _norm(registro.get("carrera", ""))
    tipo_reg = _norm(registro.get("tipo_practica", ""))
    if carrera and carrera_reg:
        puntaje += 4 if carrera_reg == _norm(carrera) else -3
    if tipo and tipo_reg:
        puntaje += 2 if tipo_reg == _norm(tipo) else 0

    return puntaje, motivo


def buscar_tutor(registros, campo, carrera=None, tipo=None, nombre_tutor=None):
    """Busca el tutor mas probable para un campo clinico.

    Devuelve una copia del registro con las claves coincidencia y revisar.
    """
    if not registros:
        return None

    campo_norm = _norm(g.limpiar_campo(campo))
    corto_norm = _norm(g.resumir_campo(campo))
    tutor_norm = _norm(nombre_tutor or g.extraer_nombre_tutor(campo))

    candidatos = []
    for registro in registros:
        puntaje, motivo = _puntaje(
            registro, campo_norm, corto_norm, tutor_norm, carrera, tipo
        )
        if puntaje:
            candidatos.append((puntaje, motivo, registro))

    if not candidatos:
        return None

    candidatos.sort(key=lambda item: item[0], reverse=True)
    mejor_puntaje, motivo, mejor = candidatos[0]
    empatados = [
        item for item in candidatos if abs(item[0] - mejor_puntaje) < 5
    ]

    resultado = dict(mejor)
    resultado["coincidencia"] = motivo
    resultado["puntaje"] = mejor_puntaje
    resultado["ambiguo"] = len(empatados) > 1
    resultado["revisar"] = "SI" if resultado["ambiguo"] or not resultado.get("correo_tutor") else ""
    return resultado


def describir(registro) -> str:
    if not registro:
        return "sin coincidencia"
    tutor = registro.get("tutor") or "sin tutor"
    correo = registro.get("correo_tutor") or "sin correo"
    return f"{tutor} <{correo}> ({registro.get('coincidencia', '')})"
