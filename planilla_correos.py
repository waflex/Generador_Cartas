#!/usr/bin/env python3
"""
planilla_correos.py
Genera una planilla Excel para correspondencia combinada.

Fuente de datos:
  - Grilla de rotaciones (campos y alumnos)
  - Listado de alumnos (RUT, correo, area, tipo)
  - Cartas generadas (para las rutas de adjuntos)

La hoja "Correspondencia" queda lista para usarse con Correo_Template.docx.
La hoja "Detalle_Alumnos" sirve para envios individuales.
"""

import html
import os
import re
import subprocess
import time
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

import generar_cartas as g
import catalogo_tutores as ct

PARTICULAS = {
    "DE", "DEL", "LA", "LAS", "LOS", "SAN", "SANTA", "VAN", "VON",
    "DA", "DI", "DOS", "DAS", "Y",
}

DIAS_SEMANA = [
    "lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo",
]

ENCABEZADO_FONDO = PatternFill("solid", fgColor="1F4E78")
ENCABEZADO_FUENTE = Font(color="FFFFFF", bold=True)
BORDE_FINO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=BORDE_FINO, right=BORDE_FINO, top=BORDE_FINO, bottom=BORDE_FINO)
ALINEACION = Alignment(wrap_text=True, vertical="top")


# ---------------------------------------------------------------------------
# CORREO INSTITUCIONAL
# ---------------------------------------------------------------------------
def _slug(texto: str) -> str:
    s = unicodedata.normalize("NFD", str(texto))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def _primer_apellido(apellidos: str) -> str:
    tokens = [t for t in re.split(r"\s+", apellidos.strip()) if t]
    if not tokens:
        return ""
    i = 0
    while i < len(tokens) - 1 and g.normalizar(tokens[i]) in PARTICULAS:
        i += 1
    return " ".join(tokens[: i + 1])


def _primer_nombre(nombres: str) -> str:
    tokens = [t for t in re.split(r"\s+", nombres.strip()) if t]
    return tokens[0] if tokens else ""


def correo_institucional(nombre: str, rut) -> str:
    """Calcula nombre.apellido#####@edu.ipchile.cl (##### = ultimos 5 del RUT sin DV)."""
    nombre = str(nombre).strip()
    if "," in nombre:
        apellidos, nombres = nombre.split(",", 1)
    else:
        partes = nombre.split()
        if len(partes) < 2:
            return ""
        apellidos, nombres = partes[-1], " ".join(partes[:-1])

    nombre_slug = _slug(_primer_nombre(nombres))
    apellido_slug = _slug(_primer_apellido(apellidos))

    digitos = re.sub(r"[^0-9Kk]", "", str(rut))
    if len(digitos) < 2:
        return ""
    cuerpo = digitos[:-1]
    cola = cuerpo[-5:].zfill(5)

    if not nombre_slug or not apellido_slug:
        return ""
    return f"{nombre_slug}.{apellido_slug}{cola}@edu.ipchile.cl"


def _correo_valido(correo: str) -> bool:
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", (correo or "").strip()))


def resolver_correo(alumno: dict) -> dict:
    """Prioriza el correo institucional del listado; calcula uno si falta."""
    listado = (alumno.get("correo") or "").strip()
    calculado = correo_institucional(
        alumno.get("nombre_original", ""), alumno.get("rut_fmt", "")
    )

    if _correo_valido(listado) and listado.lower().endswith("@edu.ipchile.cl"):
        final = listado
        revisar = "SI" if calculado and calculado.lower() != listado.lower() else ""
    elif calculado:
        final = calculado
        revisar = "SI" if _correo_valido(listado) else ""
    else:
        final = listado if _correo_valido(listado) else ""
        revisar = "SI"

    return {
        "listado": listado,
        "calculado": calculado,
        "final": final,
        "revisar": revisar,
    }


# ---------------------------------------------------------------------------
# FECHAS
# ---------------------------------------------------------------------------
def fecha_ingreso(fecha_str: str):
    """Devuelve (texto_largo, texto_corto, dia_semana) del inicio de la rotacion."""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})", fecha_str or "")
    if not m:
        return "", "", ""
    dia, mes, anio = (int(x) for x in m.groups())
    try:
        fecha = datetime(anio, mes, dia)
        dia_semana = DIAS_SEMANA[fecha.weekday()]
    except ValueError:
        dia_semana = ""
    largo = f"{dia} de {g.MESES.get(mes, '')} de {anio}".strip()
    corto = f"{dia:02d}-{mes:02d}-{anio}"
    return largo, corto, dia_semana


# ---------------------------------------------------------------------------
# TEMPLATE DE CORREO
# ---------------------------------------------------------------------------
def leer_campos_template(path: Path) -> list:
    """Lee los marcadores {{campo}} del template de correo."""
    if not path or not Path(path).exists():
        return []
    doc = Document(str(path))
    textos = [p.text for p in doc.paragraphs]
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                textos.extend(p.text for p in celda.paragraphs)
    for seccion in doc.sections:
        textos.extend(p.text for p in seccion.header.paragraphs)
        textos.extend(p.text for p in seccion.footer.paragraphs)

    campos = set()
    for texto in textos:
        for encontrado in re.findall(r"\{\{\s*(.*?)\s*\}\}", texto):
            if encontrado:
                campos.add(encontrado)
    return sorted(campos)


# ---------------------------------------------------------------------------
# CARTAS GENERADAS
# ---------------------------------------------------------------------------
def _carpeta_rotacion(rotacion: str, fecha_str: str, output_base: Path):
    fecha_safe = re.sub(r"[^\w\-]", "_", fecha_str.split(" ")[0]) if fecha_str else rotacion
    carpeta = output_base / f"{rotacion}_{fecha_safe}"
    if carpeta.exists():
        return carpeta
    if output_base.exists():
        candidatas = sorted(p for p in output_base.glob(f"{rotacion}_*") if p.is_dir())
        if candidatas:
            return candidatas[0]
    return None


def _cartas_campo(rotacion, campo, fecha_str, output_base, generados):
    campo_corto = g.resumir_campo(campo)
    campo_limpio = g.limpiar_campo(campo)
    if generados:
        rutas = []
        for item in generados:
            if item.get("rotacion") != rotacion:
                continue
            mismo_campo = (
                item.get("campo") == campo_corto
                or item.get("campo_original") == campo
                or g.normalizar(item.get("campo", "")) == g.normalizar(campo_corto)
                or g.normalizar(item.get("campo_original", "")) == g.normalizar(campo)
            )
            ruta = Path(item.get("archivo", ""))
            if mismo_campo and ruta.exists() and ruta not in rutas:
                rutas.append(ruta)
        if rutas:
            return rutas

    carpeta = _carpeta_rotacion(rotacion, fecha_str, output_base)
    if not carpeta:
        return []

    bases = []
    for texto in (campo_corto, campo_limpio, campo):
        base = f"CARTA_{rotacion}_{g.safe_filename(texto, 40)}"
        if base not in bases:
            bases.append(base)
    for base in bases:
        for extension in (".pdf", ".docx"):
            ruta = carpeta / f"{base}{extension}"
            if ruta.exists():
                return [ruta]
    return []


def _todas_las_cartas(rotacion, fecha_str, output_base):
    carpeta = _carpeta_rotacion(rotacion, fecha_str, output_base)
    if not carpeta:
        return []
    archivos = sorted(carpeta.glob(f"CARTA_{rotacion}_*.pdf"))
    if not archivos:
        archivos = sorted(carpeta.glob(f"CARTA_{rotacion}_*.docx"))
    return archivos


def _unir_rutas(rutas) -> str:
    vistas = []
    for ruta in rutas:
        texto = str(Path(ruta).resolve())
        if texto not in vistas:
            vistas.append(texto)
    return "; ".join(vistas)


# ---------------------------------------------------------------------------
# CONSTRUCCION DE DATOS
# ---------------------------------------------------------------------------
def construir_datos(grilla, maestro, rotaciones, output_base, carrera,
                    adjuntos_comunes=None, generados=None, catalogo=None,
                    tipo=None, area=None, hora_ingreso=None,
                    correo_recordatorio=None):
    adjuntos_comunes = [str(Path(a).resolve()) for a in (adjuntos_comunes or [])]
    comunes_texto = "; ".join(adjuntos_comunes)
    hora_defecto = (
        hora_ingreso
        or g.CONFIG.get("recordatorio", {}).get("hora_ingreso")
        or g.CONFIG.get("outlook", {}).get("hora_ingreso")
        or "09:00"
    )

    filas_rotacion = []
    filas_grupos = []
    filas_alumnos = []
    faltantes = []

    for rotacion in rotaciones:
        info = grilla.get(rotacion)
        if not info:
            continue

        fecha_str = info["fecha_str"]
        largo, corto, dia = fecha_ingreso(fecha_str)
        fecha_texto = g.fecha_larga(fecha_str) or largo
        campos_info = []
        grupos = []
        alumnos_rotacion = []

        for campo, nombres in info["campos"].items():
            campo_corto = g.resumir_campo(campo)
            alumnos = []
            for nombre in nombres:
                alumno, _ = g.buscar_alumno(nombre, maestro)
                if not alumno:
                    faltantes.append((rotacion, campo_corto, nombre))
                    alumnos.append({
                        "nombre_original": nombre,
                        "rut_fmt": "",
                        "correo": "",
                        "area": "",
                        "tipo": "",
                        "sin_match": True,
                    })
                else:
                    alumnos.append(alumno)

            alumnos_texto = ", ".join(
                g.nombre_visible(a["nombre_original"]) for a in alumnos
            )
            alumnos_original = ", ".join(
                a["nombre_original"] for a in alumnos
            )
            lista_alumnos = "\n".join(
                g.nombre_visible(a["nombre_original"]) for a in alumnos
            )

            correos = []
            for alumno in alumnos:
                if alumno.get("sin_match"):
                    continue
                datos = resolver_correo(alumno)
                if datos["final"] and datos["final"] not in correos:
                    correos.append(datos["final"])

            cartas = _cartas_campo(
                rotacion, campo, fecha_str, output_base, generados
            )

            tutor_data = (
                ct.buscar_tutor(catalogo, campo, carrera, tipo)
                if catalogo else None
            )
            tutor = (
                (tutor_data or {}).get("tutor")
                or g.extraer_nombre_tutor(campo)
            )
            correo_tutor = (tutor_data or {}).get("correo_tutor", "")
            direccion = (tutor_data or {}).get("direccion", "")
            contacto = (tutor_data or {}).get("contacto", "")
            nombre_residencia = (
                (tutor_data or {}).get("nombre_residencia") or campo_corto
            )
            hora = (tutor_data or {}).get("hora_ingreso") or hora_defecto
            tipo_practica = (
                (tutor_data or {}).get("tipo_practica")
                or tipo
                or g.TIPO
            )
            area_campo = (
                (tutor_data or {}).get("area")
                or area
                or g.AREA
            )
            coincidencia = (tutor_data or {}).get("coincidencia", "")
            revisar_tutor = "SI" if (not tutor or not correo_tutor) else ""
            if (tutor_data or {}).get("revisar"):
                revisar_tutor = "SI"

            asunto_presentacion = (
                f"Cartas de presentacion - {rotacion} - {campo_corto}"
            )

            item = {
                "campo": campo_corto,
                "campo_original": campo,
                "alumnos": alumnos,
                "alumnos_texto": alumnos_texto,
                "alumnos_original": alumnos_original,
                "lista_alumnos": lista_alumnos,
                "correos": correos,
                "cartas": cartas,
                "tutor": tutor,
                "correo_tutor": correo_tutor,
                "direccion": direccion,
                "contacto": contacto,
                "nombre_residencia": nombre_residencia,
                "hora_ingreso": hora,
                "tipo_practica": tipo_practica,
                "area": area_campo,
                "coincidencia_tutor": coincidencia,
                "revisar_tutor": revisar_tutor,
            }
            campos_info.append(item)
            alumnos_rotacion.append(f"{campo_corto}: {alumnos_texto}")

            grupo = {
                "Tipo": "grupo",
                "carrera": carrera,
                "fecha_ingreso": fecha_texto,
                "fecha_ingreso_corta": corto,
                "dia_ingreso": dia,
                "Rotacion": rotacion,
                "fecha_rotacion": fecha_str,
                "Campo_clinico": campo_corto,
                "Campo_clinico_Original": campo,
                "Alumnos": alumnos_texto,
                "Alumnos_rotacion": alumnos_texto,
                "Alumnos_Original": alumnos_original,
                "Lista_Alumnos": lista_alumnos,
                "Correos": "; ".join(correos),
                "Correos_Alumnos": "; ".join(correos),
                "Tutor": tutor,
                "Correo_Tutor": correo_tutor,
                "Direccion": direccion,
                "Contacto": contacto,
                "Nombre_Residencia": nombre_residencia,
                "Hora_Ingreso": hora,
                "Tipo_Practica": tipo_practica,
                "Area": area_campo,
                "Coincidencia_Tutor": coincidencia,
                "Revisar_Tutor": revisar_tutor,
                "Asunto_Presentacion": asunto_presentacion,
                "Asunto_Recordatorio": "",
                "Adjuntos_Cartas": _unir_rutas(cartas),
                "Adjuntos_Comunes": comunes_texto,
                "Total_Alumnos": len(alumnos),
                "Total_Cartas": len(cartas),
                "_campos_info": [item],
                "_alumnos_data": alumnos,
            }
            asunto_recordatorio = _asunto_desde_template(
                correo_recordatorio,
                grupo,
                f"Informacion centro de practica para el {fecha_texto} {carrera}",
            )
            grupo["Asunto_Recordatorio"] = asunto_recordatorio
            grupos.append(grupo)
            filas_grupos.append(grupo)

            for alumno in alumnos:
                datos = resolver_correo(alumno) if not alumno.get("sin_match") else {
                    "listado": "", "calculado": "", "final": "", "revisar": "SI",
                }
                filas_alumnos.append({
                    "Rotacion": rotacion,
                    "fecha_rotacion": fecha_str,
                    "fecha_ingreso": fecha_texto,
                    "carrera": carrera,
                    "Tipo_Practica": tipo_practica,
                    "Area_Campo": area_campo,
                    "Campo_clinico": campo_corto,
                    "Campo_clinico_Original": campo,
                    "Alumnos_rotacion": alumnos_texto,
                    "Nombre": alumno.get("nombre_original", ""),
                    "RUT": alumno.get("rut_fmt", ""),
                    "Area": alumno.get("area") or area_campo,
                    "Tipo": alumno.get("tipo") or tipo_practica,
                    "Tutor": tutor,
                    "Correo_Tutor": correo_tutor,
                    "Direccion": direccion,
                    "Hora_Ingreso": hora,
                    "Revisar_Tutor": revisar_tutor,
                    "Correo_Listado": datos["listado"],
                    "Correo_Calculado": datos["calculado"],
                    "Correo_Final": datos["final"],
                    "Revisar_Correo": datos["revisar"],
                    "Adjunto_Carta": _unir_rutas(cartas),
                    "Adjuntos_Comunes": comunes_texto,
                    "Asunto": asunto_presentacion,
                    "Asunto_Recordatorio": asunto_recordatorio,
                    "_campos_info": [item],
                })

        if not campos_info:
            continue

        cartas_rotacion = []
        for item in campos_info:
            cartas_rotacion.extend(item["cartas"])
        if not cartas_rotacion:
            cartas_rotacion = _todas_las_cartas(rotacion, fecha_str, output_base)

        correos_rotacion = []
        for item in campos_info:
            for correo in item["correos"]:
                if correo not in correos_rotacion:
                    correos_rotacion.append(correo)

        asunto = f"Cartas de presentacion - {rotacion} - {carrera}"

        fila = {
            "carrera": carrera,
            "fecha_ingreso": fecha_texto,
            "fecha_ingreso_corta": corto,
            "dia_ingreso": dia,
            "Rotacion": rotacion,
            "fecha_rotacion": fecha_str,
            "Campo_clinico": "\n".join(item["campo"] for item in campos_info),
            "Alumnos_rotacion": "\n".join(alumnos_rotacion),
            "Correos": "; ".join(correos_rotacion),
            "Asunto": asunto,
            "Adjuntos_Cartas": _unir_rutas(cartas_rotacion),
            "Adjuntos_Comunes": comunes_texto,
            "Total_Campos": len(campos_info),
            "Total_Grupos": len(grupos),
            "Total_Alumnos": sum(len(item["alumnos"]) for item in campos_info),
            "Total_Cartas": len(cartas_rotacion),
            "_campos_info": campos_info,
            "_grupos": grupos,
        }
        for indice, item in enumerate(campos_info, start=1):
            fila[f"Campo_clinico_{indice}"] = item["campo"]
            fila[f"Alumnos_{indice}"] = item["alumnos_texto"]
            fila[f"Correos_{indice}"] = "; ".join(item["correos"])
            fila[f"Adjuntos_{indice}"] = _unir_rutas(item["cartas"])
            fila[f"Tutor_{indice}"] = item["tutor"]
            fila[f"Correo_Tutor_{indice}"] = item["correo_tutor"]
        filas_rotacion.append(fila)

    return filas_rotacion, filas_grupos, filas_alumnos, faltantes, adjuntos_comunes


# ---------------------------------------------------------------------------
# ESCRITURA DEL EXCEL
# ---------------------------------------------------------------------------
def _estilo_encabezado(ws, n_columnas: int):
    for columna in range(1, n_columnas + 1):
        celda = ws.cell(row=1, column=columna)
        celda.fill = ENCABEZADO_FONDO
        celda.font = ENCABEZADO_FUENTE
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = BORDE
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"


def _escribir_hoja(ws, encabezados, filas, ajustar_ancho=True):
    ws.append(encabezados)
    for fila in filas:
        ws.append([fila.get(col, "") for col in encabezados])
    _estilo_encabezado(ws, len(encabezados))

    for fila in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for celda in fila:
            celda.alignment = ALINEACION
            celda.border = BORDE

    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(encabezados))}{ws.max_row}"

    if ajustar_ancho:
        for indice, columna in enumerate(encabezados, start=1):
            largo = len(str(columna))
            for celda in ws[get_column_letter(indice)]:
                if celda.value is not None:
                    largo = max(largo, len(str(celda.value)))
            ws.column_dimensions[get_column_letter(indice)].width = min(max(12, largo + 2), 60)


def _agregar_hipervinculos(ws, encabezados, nombre_columna: str):
    """Convierte rutas de una sola columna en enlaces clicables."""
    if nombre_columna not in encabezados:
        return
    columna = encabezados.index(nombre_columna) + 1
    for fila in range(2, ws.max_row + 1):
        celda = ws.cell(row=fila, column=columna)
        if not celda.value:
            continue
        primera = str(celda.value).split(";")[0].strip()
        if Path(primera).exists():
            celda.hyperlink = primera
            celda.font = Font(color="0563C1", underline="single")


def exportar_planilla(grilla, maestro, rotaciones, output_base: Path,
                      destino: Path, carrera: str,
                      adjuntos_comunes=None, generados=None,
                      correo_template: Path = None,
                      correo_recordatorio: Path = None,
                      catalogo=None, tipo=None, area=None,
                      hora_ingreso=None):
    (
        filas_rotacion, filas_grupos, filas_alumnos, faltantes, comunes
    ) = construir_datos(
        grilla, maestro, rotaciones, output_base, carrera,
        adjuntos_comunes, generados, catalogo=catalogo, tipo=tipo,
        area=area, hora_ingreso=hora_ingreso,
        correo_recordatorio=correo_recordatorio,
    )

    libro = openpyxl.Workbook()

    ws = libro.active
    ws.title = "Correspondencia"
    encabezados_rotacion = [
        "carrera", "fecha_ingreso", "fecha_ingreso_corta", "dia_ingreso",
        "Rotacion", "fecha_rotacion", "Campo_clinico", "Alumnos_rotacion",
        "Correos", "Asunto", "Adjuntos_Cartas", "Adjuntos_Comunes",
        "Total_Campos", "Total_Grupos", "Total_Alumnos", "Total_Cartas",
    ]
    max_campos = max(
        (len([k for k in fila if k.startswith("Campo_clinico_")]) for fila in filas_rotacion),
        default=0,
    )
    for indice in range(1, max_campos + 1):
        encabezados_rotacion.extend([
            f"Campo_clinico_{indice}", f"Alumnos_{indice}",
            f"Correos_{indice}", f"Adjuntos_{indice}",
            f"Tutor_{indice}", f"Correo_Tutor_{indice}",
        ])
    _escribir_hoja(ws, encabezados_rotacion, filas_rotacion)
    _agregar_hipervinculos(ws, encabezados_rotacion, "Adjuntos_Cartas")

    ws_grupos = libro.create_sheet("Correspondencia_Grupos")
    encabezados_grupos = [
        "carrera", "Tipo_Practica", "Area", "Rotacion", "fecha_rotacion",
        "fecha_ingreso", "fecha_ingreso_corta", "dia_ingreso",
        "Campo_clinico", "Campo_clinico_Original", "Alumnos",
        "Lista_Alumnos", "Correos", "Tutor", "Correo_Tutor", "Direccion",
        "Nombre_Residencia", "Hora_Ingreso", "Contacto",
        "Asunto_Presentacion", "Asunto_Recordatorio",
        "Adjuntos_Cartas", "Adjuntos_Comunes", "Total_Alumnos",
        "Total_Cartas", "Revisar_Tutor", "Coincidencia_Tutor",
    ]
    _escribir_hoja(ws_grupos, encabezados_grupos, filas_grupos)
    _agregar_hipervinculos(ws_grupos, encabezados_grupos, "Adjuntos_Cartas")

    ws_detalle = libro.create_sheet("Detalle_Alumnos")
    encabezados_alumnos = [
        "Rotacion", "fecha_rotacion", "fecha_ingreso", "carrera",
        "Tipo_Practica", "Area_Campo", "Campo_clinico",
        "Campo_clinico_Original", "Alumnos_rotacion", "Nombre", "RUT",
        "Area", "Tipo", "Tutor", "Correo_Tutor", "Direccion",
        "Hora_Ingreso", "Revisar_Tutor", "Correo_Listado",
        "Correo_Calculado", "Correo_Final", "Revisar_Correo",
        "Adjunto_Carta", "Adjuntos_Comunes", "Asunto",
        "Asunto_Recordatorio",
    ]
    _escribir_hoja(ws_detalle, encabezados_alumnos, filas_alumnos)
    _agregar_hipervinculos(ws_detalle, encabezados_alumnos, "Adjunto_Carta")

    ws_comunes = libro.create_sheet("Adjuntos_Comunes")
    filas_comunes = [
        {"Archivo": Path(ruta).name, "Ruta": ruta, "Existe": "SI" if Path(ruta).exists() else "NO"}
        for ruta in comunes
    ]
    _escribir_hoja(ws_comunes, ["Archivo", "Ruta", "Existe"], filas_comunes)

    ws_instrucciones = libro.create_sheet("Instrucciones", 0)
    instrucciones = [
        {"Paso": "1", "Indicacion": "Los borradores se crean directamente desde la interfaz, uno por grupo y por plantilla."},
        {"Paso": "2", "Indicacion": "Use la hoja Correspondencia_Grupos para revisar un registro por campo clinico y sus tutores."},
        {"Paso": "3", "Indicacion": "En la hoja Correspondencia_Grupos, la columna Correos contiene a los alumnos del grupo y Correo_Tutor al tutor en copia."},
        {"Paso": "4", "Indicacion": "Revise Revisar_Tutor = SI: falta correo de tutor o el catalogo no fue concluyente."},
        {"Paso": "5", "Indicacion": "El correo_tutor se obtiene del tercer Excel: columna Correo_Tutor si existe; si no, desde Tutor(a) o Contacto."},
        {"Paso": "6", "Indicacion": "Para envios individuales use la hoja Detalle_Alumnos."},
        {"Paso": "7", "Indicacion": "Word no adjunta archivos por destinatario: use Adjuntos_Cartas y Adjuntos_Comunes para ubicarlos."},
        {"Paso": "8", "Indicacion": "Revise las filas con Revisar_Correo = SI antes de enviar."},
    ]
    _escribir_hoja(ws_instrucciones, ["Paso", "Indicacion"], instrucciones)
    ws_instrucciones.column_dimensions["A"].width = 8
    ws_instrucciones.column_dimensions["B"].width = 110

    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    libro.save(destino)

    campos_template = []
    for plantilla in (correo_template, correo_recordatorio):
        if not plantilla:
            continue
        for campo in leer_campos_template(plantilla):
            if campo not in campos_template:
                campos_template.append(campo)

    disponibles = {
        _clave_placeholder(col)
        for col in (
            encabezados_rotacion + encabezados_grupos + encabezados_alumnos
        )
    }
    alias_placeholder = {
        "DIRECCION RESIDENCIA": "DIRECCION",
        "NOMBRE TUTOR": "TUTOR",
    }
    faltantes_template = [
        campo for campo in campos_template
        if _clave_placeholder(campo) not in disponibles
        and alias_placeholder.get(_clave_placeholder(campo)) not in disponibles
    ]

    return {
        "archivo": destino,
        "rotaciones": len(filas_rotacion),
        "grupos": len(filas_grupos),
        "alumnos": len(filas_alumnos),
        "rotaciones_sin_adjuntos": sum(
            1 for fila in filas_rotacion if not fila.get("Adjuntos_Cartas")
        ),
        "grupos_sin_adjuntos": sum(
            1 for fila in filas_grupos if not fila.get("Adjuntos_Cartas")
        ),
        "grupos_sin_tutor": sum(
            1 for fila in filas_grupos if fila.get("Revisar_Tutor")
        ),
        "tutores_ambiguos": sum(
            1 for fila in filas_grupos
            if fila.get("Revisar_Tutor") and fila.get("Coincidencia_Tutor")
        ),
        "faltantes": faltantes,
        "campos_template": campos_template,
        "campos_faltantes_template": faltantes_template,
    }


# ---------------------------------------------------------------------------
# BORRADORES DE OUTLOOK
# ---------------------------------------------------------------------------
def _texto_html(valor) -> str:
    return html.escape(str(valor or "")).replace("\n", "<br>")


def _valores_correo(fila: dict) -> dict:
    """Valores disponibles para los placeholders de ambos templates."""
    campo = fila.get("Campo_clinico") or fila.get("Campo_Clinico") or ""
    alumnos = fila.get("Alumnos") or fila.get("Alumnos_rotacion") or ""
    lista = fila.get("Lista_Alumnos") or alumnos.replace(", ", "\n")
    return {
        "carrera": fila.get("carrera", ""),
        "Carrera": fila.get("carrera", ""),
        "fecha_ingreso": fila.get("fecha_ingreso", ""),
        "Fecha_Ingreso": fila.get("fecha_ingreso", ""),
        "fecha_ingreso_corta": fila.get("fecha_ingreso_corta", ""),
        "dia_ingreso": fila.get("dia_ingreso", ""),
        "Rotacion": fila.get("Rotacion", ""),
        "fecha_rotacion": fila.get("fecha_rotacion", ""),
        "Campo_clinico": campo,
        "Campo_Clinico": campo,
        "Campo_clinico_Original": fila.get("Campo_clinico_Original", ""),
        "Alumnos": alumnos,
        "Alumnos_rotacion": alumnos,
        "Lista_Alumnos": lista,
        "Tutor": fila.get("Tutor", ""),
        "Nombre_Tutor": fila.get("Tutor", ""),
        "Correo_Tutor": fila.get("Correo_Tutor", ""),
        "Direccion": fila.get("Direccion", ""),
        "Direccion_Residencia": fila.get("Direccion", ""),
        "Nombre_residencia": fila.get("Nombre_Residencia", campo),
        "Nombre_Residencia": fila.get("Nombre_Residencia", campo),
        "Hora_Ingreso": fila.get("Hora_Ingreso", ""),
        "Tipo_Practica": fila.get("Tipo_Practica", ""),
        "Area": fila.get("Area", ""),
        "Contacto": fila.get("Contacto", ""),
        "Asunto_Presentacion": fila.get("Asunto_Presentacion", ""),
        "Asunto_Recordatorio": fila.get("Asunto_Recordatorio", ""),
    }


def _clave_placeholder(texto: str) -> str:
    return g.normalizar(texto).replace("_", " ").strip()


def _reemplazar_marcadores(texto: str, valores: dict) -> str:
    valores_norm = {
        _clave_placeholder(clave): valor for clave, valor in valores.items()
    }

    def _reemplazo(match):
        clave = _clave_placeholder(match.group(1))
        if clave not in valores_norm:
            return match.group(0)
        return _texto_html(valores_norm[clave])

    return re.sub(r"\{\{\s*(.*?)\s*\}\}", _reemplazo, texto)




def _tabla_html(fila: dict) -> str:
    filas = [
        '<tr><td style="width: 35%;"><b>Campo Clinico</b></td>'
        f'<td>{_texto_html(fila.get("fecha_rotacion"))}</td></tr>',
    ]

    for item in fila.get("_campos_info", []):
        alumnos = [
            alumno.strip()
            for alumno in (item.get("alumnos_texto") or "").split(",")
            if alumno.strip()
        ]

        if not alumnos:
            alumnos = [""]

        for i, alumno in enumerate(alumnos):
            if i == 0:
                filas.append(
                    f'<tr><td rowspan="{len(alumnos)}" '
                    f'style="vertical-align: middle;">'
                    f'{_texto_html(item.get("campo"))}</td>'
                    f'<td>{_texto_html(alumno)}</td></tr>'
                )
            else:
                filas.append(
                    f'<tr><td>{_texto_html(alumno)}</td></tr>'
                )

    return (
        '<table border="1" cellpadding="6" cellspacing="0" '
        'style="border-collapse: collapse; width: 100%; font-size: 10.5pt;">'
        + "".join(filas)
        + "</table>"
    )




def _template_docx(correo_template):
    if not correo_template:
        return None
    try:
        ruta = g.asegurar_template_docx(correo_template, log=lambda *_: None)
    except Exception:
        return None
    return ruta if ruta and Path(ruta).exists() else None


def _asunto_desde_template(correo_template, fila: dict, por_defecto: str = "") -> str:
    """Lee una linea 'Asunto:' del template y la convierte en asunto real."""
    ruta = _template_docx(correo_template)
    if not ruta:
        return por_defecto

    doc = Document(str(ruta))
    valores = _valores_correo(fila)
    for parrafo in doc.paragraphs:
        texto = parrafo.text.strip()
        if not texto:
            continue
        if re.match(r"^asunto\s*:", texto, flags=re.IGNORECASE):
            asunto = re.sub(
                r"^asunto\s*:\s*", "", texto, flags=re.IGNORECASE
            )
            asunto = _reemplazar_marcadores(asunto, valores)
            asunto = html.unescape(asunto).replace("<br>", " ")
            return re.sub(r"\s+", " ", asunto).strip() or por_defecto
        break
    return por_defecto


def construir_cuerpo_html(correo_template, fila: dict, tipo_correo="presentacion") -> str:
    """Arma el cuerpo HTML de presentacion o recordatorio."""
    valores = _valores_correo(fila)
    partes = []
    ruta = _template_docx(correo_template)

    if ruta:
        doc = Document(str(ruta))
        for hijo in doc.element.body.iterchildren():
            if hijo.tag == qn("w:p"):
                parrafo = Paragraph(hijo, doc)
                texto_original = parrafo.text.strip()
                if (
                    tipo_correo == "recordatorio"
                    and re.match(r"^asunto\s*:", texto_original, flags=re.IGNORECASE)
                ):
                    continue
                texto = _reemplazar_marcadores(parrafo.text, valores)
                if texto.strip():
                    partes.append(f"<p>{texto}</p>")
                else:
                    partes.append("<p>&nbsp;</p>")
            elif hijo.tag == qn("w:tbl"):
                continue
    else:
        if tipo_correo == "recordatorio":
            partes.extend([
                f"<p>Estimada(o) estudiantes de {_texto_html(valores['Carrera'])}:</p>",
                f"<p>{_texto_html(valores['Lista_Alumnos'])}</p>",
                "<p>Junto con saludar, se recuerda el inicio de su internado en "
                f"{_texto_html(valores['Nombre_residencia'])}, "
                f"ubicado en {_texto_html(valores['Direccion_Residencia'])}.</p>",
                f"<p>Se presentan con {_texto_html(valores['Nombre_Tutor'])} "
                f"a las {_texto_html(valores['Hora_Ingreso'])} horas.</p>",
            ])
        else:
            partes.extend([
                "<p>Estimados estudiantes de preclinica de "
                f"{_texto_html(valores['carrera'])}, por el presente adjunto carta "
                "de presentacion quienes ingresa este lunes "
                f"{_texto_html(valores['fecha_ingreso'])} respectivamente.</p>",
                "<p>&nbsp;</p>",
                "<p>Tambien se adjunta seguro escolar con su esquema de uso y "
                "registro de asistencia.</p>",
                "<p>&nbsp;</p>",
                "<p>Favor de leer las instrucciones de ingreso para cada campo "
                "clinico actualizado.</p>",
                "<p>&nbsp;</p>",
            ])

    if tipo_correo != "recordatorio":
        partes.append(_tabla_html(fila))

    return (
        '<html><body style="font-family: Calibri, Arial, sans-serif; '
        'font-size: 11pt; color: #000000;">'
        + "".join(partes)
        + "</body></html>"
    )


def _rutas_adjuntos(texto) -> list:
    rutas = []
    for parte in str(texto or "").split(";"):
        ruta = parte.strip()
        if ruta and Path(ruta).exists() and ruta not in rutas:
            rutas.append(ruta)
    return rutas


def _ruta_outlook():
    candidatos = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
        / "Microsoft Office" / "root" / "Office16" / "OUTLOOK.EXE",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
        / "Microsoft Office" / "root" / "Office16" / "OUTLOOK.EXE",
    ]
    for ruta in candidatos:
        if ruta.exists():
            return ruta
    return None


def _obtener_outlook():
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "Falta pywin32 para crear borradores en Outlook. "
            "Instala con: pip install pywin32"
        ) from exc

    try:
        return win32.Dispatch("Outlook.Application")
    except Exception:
        pass

    ruta = _ruta_outlook()
    if not ruta:
        raise RuntimeError(
            "No se pudo iniciar Outlook. Abrelo manualmente e intenta de nuevo."
        )
    try:
        subprocess.Popen([str(ruta)])
    except Exception as exc:
        raise RuntimeError(f"No se pudo iniciar Outlook: {exc}") from exc

    for _ in range(6):
        time.sleep(3)
        try:
            return win32.Dispatch("Outlook.Application")
        except Exception:
            continue
    raise RuntimeError(
        "Outlook no respondio. Abrelo manualmente e intenta de nuevo."
    )


def _lista_correos(texto) -> list:
    correos = []
    for parte in str(texto or "").split(";"):
        correo = parte.strip()
        if correo and correo.lower() not in [c.lower() for c in correos]:
            correos.append(correo)
    return correos


def _fila_para_alumno(fila: dict, alumno: dict) -> dict:
    copia = dict(fila)
    nombre = g.nombre_visible(alumno.get("Nombre", ""))
    copia["Alumnos"] = nombre
    copia["Alumnos_rotacion"] = nombre
    copia["Lista_Alumnos"] = nombre
    copia["_campos_info"] = []
    for item in fila.get("_campos_info", []):
        nuevo = dict(item)
        nuevo["alumnos_texto"] = nombre
        nuevo["alumnos"] = [alumno]
        copia["_campos_info"].append(nuevo)
    return copia


def crear_borradores_outlook(
    grilla, maestro, rotaciones, output_base: Path, carrera: str,
    adjuntos_comunes=None, generados=None, correo_template=None,
    correo_recordatorio=None, catalogo=None,
    modo="grupo_presentacion_recordatorio", usar_cco=False,
    cc_tutor=True, mostrar=False, tipo=None, area=None,
    hora_ingreso=None, por_alumno=False, log=None,
):
    """Crea borradores en Outlook. No envia correos."""
    log = log or (lambda *_: None)

    if por_alumno:
        modo = "alumno_presentacion"

    filas_rotacion, filas_grupos, filas_alumnos, faltantes, _ = construir_datos(
        grilla, maestro, rotaciones, output_base, carrera,
        adjuntos_comunes, generados, catalogo=catalogo, tipo=tipo,
        area=area, hora_ingreso=hora_ingreso,
        correo_recordatorio=correo_recordatorio,
    )
    filas_por_rotacion = {fila["Rotacion"]: fila for fila in filas_rotacion}
    grupos_por_clave = {
        (fila.get("Rotacion", ""), fila.get("Campo_clinico", "")): fila
        for fila in filas_grupos
    }

    outlook = _obtener_outlook()

    creados = []
    sin_correo = []
    sin_adjuntos = []
    sin_tutor = []
    errores = []
    mostrado = False
    adjuntos_reportados = set()

    def _crear(destinatarios, asunto, cuerpo, rutas, cco=None, cc=None):
        nonlocal mostrado
        mail = outlook.CreateItem(0)
        if cco is None:
            cco = usar_cco
        if cco:
            mail.BCC = "; ".join(destinatarios)
        else:
            mail.To = "; ".join(destinatarios)
        if cc:
            mail.CC = "; ".join(cc)
        mail.Subject = asunto
        mail.HTMLBody = cuerpo

        adjuntos = []
        for ruta in rutas:
            if ruta not in adjuntos:
                adjuntos.append(ruta)
                mail.Attachments.Add(ruta)
        mail.Save()
        if mostrar and not mostrado:
            mail.Display()
            mostrado = True
        return adjuntos

    def _crear_plantilla(fila, tipo_correo, destinatarios, etiqueta,
                         alumno=None):
        if not destinatarios:
            if etiqueta not in sin_correo:
                sin_correo.append(etiqueta)
            return
        if tipo_correo == "recordatorio" and not correo_recordatorio:
            return
        if tipo_correo == "recordatorio" and not Path(correo_recordatorio).exists():
            log(f"Sin plantilla de recordatorio valida: {correo_recordatorio}")
            return

        fila_correo = _fila_para_alumno(fila, alumno) if alumno else fila
        if tipo_correo == "recordatorio":
            asunto = fila_correo.get("Asunto_Recordatorio") or _asunto_desde_template(
                correo_recordatorio,
                fila_correo,
                f"Informacion centro de practica para el "
                f"{fila_correo.get('fecha_ingreso', '')} "
                f"{fila_correo.get('carrera', '')}",
            )
            cuerpo = construir_cuerpo_html(
                correo_recordatorio, fila_correo, "recordatorio"
            )
            plantilla = correo_recordatorio
        else:
            asunto = fila_correo.get("Asunto_Presentacion") or fila_correo.get(
                "Asunto", ""
            )
            cuerpo = construir_cuerpo_html(
                correo_template, fila_correo, "presentacion"
            )
            plantilla = correo_template

        rutas = _rutas_adjuntos(fila_correo.get("Adjuntos_Cartas"))
        rutas += _rutas_adjuntos(fila_correo.get("Adjuntos_Comunes"))
        if not rutas:
            if etiqueta not in adjuntos_reportados:
                adjuntos_reportados.add(etiqueta)
                sin_adjuntos.append(etiqueta)

        cc = []
        if cc_tutor:
            cc = _lista_correos(fila_correo.get("Correo_Tutor"))
            if not cc and etiqueta not in sin_tutor:
                sin_tutor.append(etiqueta)

        try:
            adjuntos = _crear(
                destinatarios, asunto, cuerpo, rutas, cc=cc
            )
            creados.append({
                "tipo": tipo_correo,
                "rotacion": fila_correo.get("Rotacion", ""),
                "campo": fila_correo.get("Campo_clinico", ""),
                "nombre": alumno.get("Nombre", "") if alumno else "",
                "destinatarios": destinatarios,
                "cc": cc,
                "asunto": asunto,
                "adjuntos": adjuntos,
                "etiqueta": etiqueta,
                "plantilla": str(plantilla or ""),
            })
            log(
                f"Borrador {tipo_correo}: {etiqueta} "
                f"({len(destinatarios)} alumnos, {len(adjuntos)} adjuntos"
                + (f", CC {len(cc)}" if cc else "")
                + ")"
            )
        except Exception as exc:
            errores.append((f"{etiqueta} [{tipo_correo}]", str(exc)))
            log(f"ERROR creando {etiqueta} [{tipo_correo}]: {exc}")

    incluye_presentacion = modo in (
        "grupo_presentacion_recordatorio", "grupo_presentacion",
        "alumno_presentacion", "rotacion_presentacion",
    )
    incluye_recordatorio = modo in (
        "grupo_presentacion_recordatorio", "grupo_recordatorio",
    )

    if modo.startswith("grupo"):
        if incluye_recordatorio and not correo_recordatorio:
            log("No se indico plantilla de recordatorio. Se omite ese borrador.")
        for fila in filas_grupos:
            destinatarios = _lista_correos(fila.get("Correos_Alumnos"))
            etiqueta = (
                f"{fila.get('Rotacion', '')} - "
                f"{fila.get('Campo_clinico', '')}"
            )
            if incluye_presentacion:
                _crear_plantilla(
                    fila, "presentacion", destinatarios, etiqueta
                )
            if incluye_recordatorio:
                _crear_plantilla(
                    fila, "recordatorio", destinatarios, etiqueta
                )
    elif modo == "alumno_presentacion":
        for alumno in filas_alumnos:
            correo = (alumno.get("Correo_Final") or "").strip()
            clave = (
                alumno.get("Rotacion", ""),
                alumno.get("Campo_clinico", ""),
            )
            fila = grupos_por_clave.get(clave) or filas_por_rotacion.get(
                alumno.get("Rotacion")
            )
            if not fila:
                continue
            etiqueta = (
                f"{alumno.get('Rotacion', '')} - "
                f"{alumno.get('Nombre', '')} - "
                f"{alumno.get('Campo_clinico', '')}"
            )
            _crear_plantilla(
                fila, "presentacion", [correo] if correo else [],
                etiqueta, alumno=alumno,
            )
    else:
        for fila in filas_rotacion:
            destinatarios = _lista_correos(fila.get("Correos"))
            etiqueta = fila.get("Rotacion", "")
            _crear_plantilla(
                fila, "presentacion", destinatarios, etiqueta
            )

    return {
        "borradores": creados,
        "creados": len(creados),
        "sin_correo": sin_correo,
        "sin_adjuntos": sin_adjuntos,
        "sin_tutor": sin_tutor,
        "errores": errores,
        "faltantes": faltantes,
        "grupos": len(filas_grupos),
        "total_borradores": len(creados),
    }


# ---------------------------------------------------------------------------
# PRUEBA DIRECTA
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    base = Path(__file__).resolve().parent
    maestro = g.cargar_maestro(base / "Lista_Preclinicos_II_Kinesiologia.xlsx")
    grilla = g.parsear_grilla(
        base / "Lista_Preclinicos_II_Kinesiologia_para_compartir_2.xlsx"
    )
    rotaciones = g.ordenar_rotaciones(grilla.keys())
    resultado = exportar_planilla(
        grilla, maestro, rotaciones, base / "cartas_output",
        base / "cartas_output" / "Planilla_Correos.xlsx", g.CARRERA,
        correo_template=base / "Correo_Template.docx",
        correo_recordatorio=base / "Estimada Estudiante recordatorio.docx",
    )
    print(resultado)
