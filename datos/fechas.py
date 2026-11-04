"""Utilidades para interpretar fechas de las planillas de rotaciones."""

import re
from datetime import date, datetime

RE_FECHA = re.compile(
    r"\b(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})[-/](\d{1,2})[-/](\d{4}))\b"
)


def _fecha(valor):
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return None


def _desde_match(match):
    if match.group(1):
        anio, mes, dia = (int(match.group(i)) for i in (1, 2, 3))
    else:
        dia, mes, anio = (int(match.group(i)) for i in (4, 5, 6))
    try:
        return date(anio, mes, dia)
    except ValueError:
        return None


def fechas_en_texto(valor):
    """Devuelve todas las fechas encontradas, en el orden del texto."""
    fecha = _fecha(valor)
    if fecha:
        return [fecha]
    texto = str(valor or "")
    resultado = []
    for match in RE_FECHA.finditer(texto):
        fecha = _desde_match(match)
        if fecha and fecha not in resultado:
            resultado.append(fecha)
    return resultado


def iso(fecha):
    return fecha.isoformat() if fecha else ""


def formato_largo(fecha):
    if not fecha:
        return ""
    return fecha.strftime("%d-%m-%Y")


def parsear_rango_fechas(valor):
    """Convierte una celda de fechas en un rango normalizado.

    Acepta:
      - 31-08-2026 al 04-09-2026
      - 25-09-2026 28-10-2026
      - 2026-09-25 00:00:00
      - objetos date/datetime de Excel
    """
    texto = "" if valor is None else str(valor).strip()
    fechas = fechas_en_texto(valor)
    inicio = fechas[0] if fechas else None
    fin = fechas[1] if len(fechas) > 1 else None

    if inicio and fin:
        normalizado = f"{formato_largo(inicio)} al {formato_largo(fin)}"
    elif inicio:
        normalizado = formato_largo(inicio)
    else:
        normalizado = texto

    return {
        "fecha_inicio": iso(inicio),
        "fecha_fin": iso(fin),
        "fecha_texto": texto,
        "fecha_texto_normalizado": normalizado,
    }


def sugerir_periodo(fecha_inicio, fecha_fin=None):
    """Sugiere un periodo academico simple, por ejemplo 2026-S2."""
    inicio = fechas_en_texto(fecha_inicio)
    fin = fechas_en_texto(fecha_fin)
    referencia = inicio[0] if inicio else (fin[0] if fin else None)
    if not referencia:
        referencia = date.today()
    semestre = "S1" if referencia.month <= 6 else "S2"
    return f"{referencia.year}-{semestre}"
