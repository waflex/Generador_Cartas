#!/usr/bin/env python3
"""
generar_cartas.py
Genera una carta PDF por campo clinico/rotacion con todos sus alumnos en la tabla.
Usa un template .doc o .docx con variables {{...}} y clona filas dinamicamente.

Dependencias:
    pip install python-docx openpyxl docx2pdf pywin32

Uso CLI:
    python generar_cartas.py --rotacion R9
    python generar_cartas.py --rotacion R9 R10
    python generar_cartas.py --all
    python generar_cartas.py --rotacion R9 --solo-docx

Uso con interfaz:
    python interfaz_cartas.py
"""

import argparse
import contextlib
import copy
import io
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import openpyxl
from docx import Document

import config as app_config
from datos.fechas import parsear_rango_fechas

# ---------------------------------------------------------------------------
# RUTAS POR DEFECTO - junto al exe/script o dentro del paquete
# ---------------------------------------------------------------------------
BASE_DIR = app_config.app_dir()
RESOURCE_DIR = app_config.resource_dir()
TEMPLATE_DOCX = app_config.ruta_recurso("carta_template.docx")
TEMPLATE_DOC = app_config.ruta_recurso("carta_template.doc")
MAESTRO_XLSX = app_config.ruta_recurso("Lista_Preclinicos_II_Kinesiologia.xlsx")
GRILLA_XLSX = app_config.ruta_recurso(
    "Lista_Preclinicos_II_Kinesiologia_para_compartir_2.xlsx"
)
OUTPUT_DIR = BASE_DIR / "cartas_output"
CONFIG = app_config.cargar_config()
CAMPOS_CLINICOS = CONFIG.get("campos_clinicos", {})
ALIAS_CAMPOS = CONFIG.get("alias_campos", {})

# ---------------------------------------------------------------------------
# CONFIGURACION INSTITUCIONAL POR DEFECTO
# ---------------------------------------------------------------------------
CARRERA = "KINESIOLOGIA"
TIPO = "practica/preclinica"
AREA = "Preclinica II"
MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre",
    11: "noviembre", 12: "diciembre",
}

IGNORAR_CELDAS = {
    "no disponible tutor en campo clinico",
    "no disponible", "\xa0", " ", "",
}

HOJA_MAESTRO = "PRECLINICA II"

PALABRAS_CLAVE_CAMPO = [
    "musculo", "esqueletico", "cardio", "neuro", "cedam", "hospital",
    "araucaria", "maskine", "avila", "roa", "kinesica",
]

TOKENS_NO_ALUMNO = {
    "desde", "semana", "horario", "tutor", "lunes", "martes", "miercoles",
    "jueves", "viernes", "sabado", "domingo", "lun", "mar", "mie", "jue",
    "vie", "sab", "dom", "ene", "feb", "abr", "may", "jun", "jul", "ago",
    "sep", "oct", "nov", "dic",
}

TOKENS_NOTA_CAMPO = {
    "lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo",
    "lun", "mie", "jue", "vie", "sab", "dom", "horario", "semana",
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def normalizar(texto: str) -> str:
    if not texto:
        return ""
    s = str(texto).strip().upper()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def formatear_rut(rut_raw) -> str:
    if isinstance(rut_raw, float) and rut_raw.is_integer():
        rut_raw = int(rut_raw)
    rut = str(rut_raw).strip().upper().replace(".", "").replace("-", "")
    if len(rut) < 2:
        return str(rut_raw)
    cuerpo, dv = rut[:-1], rut[-1]
    fmt = ""
    for i, c in enumerate(reversed(cuerpo)):
        if i > 0 and i % 3 == 0:
            fmt = "." + fmt
        fmt = c + fmt
    return f"{fmt}-{dv}"


def safe_filename(texto: str, maxlen: int = 50) -> str:
    s = unicodedata.normalize("NFD", str(texto)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^\w\s\-]", "", s)
    return re.sub(r"\s+", "_", s).strip("_")[:maxlen]


def limpiar_campo(campo: str) -> str:
    """Quita notas entre parentesis y horarios anexados al nombre del campo."""
    s = re.sub(r"\(.*?\)", "", str(campo))
    s = re.split(r"\bdesde\b", s, flags=re.IGNORECASE)[0]
    return re.sub(r"\s+", " ", s).strip()


def resumir_campo(campo: str) -> str:
    """Aplica los nombres cortos configurados en config.json."""
    original = str(campo or "")
    limpio = limpiar_campo(original)

    for alias, objetivo in sorted(
        ALIAS_CAMPOS.items(), key=lambda item: len(item[0]), reverse=True
    ):
        if normalizar(limpio) == normalizar(alias):
            limpio = str(objetivo).strip()
            break

    norm = normalizar(limpio)
    for nombre, resumen in sorted(
        CAMPOS_CLINICOS.items(), key=lambda item: len(item[0]), reverse=True
    ):
        clave = normalizar(nombre)
        if not clave:
            continue
        patron = rf"(?<!\w){re.escape(clave)}(?!\w)"
        if re.search(patron, norm):
            return str(resumen).strip()

    return re.sub(r"\s+", " ", limpio).strip()


def recargar_config():
    """Recarga config.json en caliente (util para pruebas o cambios manuales)."""
    global CONFIG, CAMPOS_CLINICOS, ALIAS_CAMPOS
    CONFIG = app_config.cargar_config()
    CAMPOS_CLINICOS = CONFIG.get("campos_clinicos", {})
    ALIAS_CAMPOS = CONFIG.get("alias_campos", {})
    return CONFIG


def fecha_larga(fecha_str: str) -> str:
    """Convierte 02-11-2026 en 02 de noviembre de 2026."""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})", str(fecha_str or ""))
    if not m:
        return str(fecha_str or "")
    dia, mes, anio = (int(x) for x in m.groups())
    return f"{dia:02d} de {MESES.get(mes, '')} de {anio}"


def nombre_visible(nombre: str) -> str:
    """Devuelve NOMBRES APELLIDOS para listas de correo."""
    texto = str(nombre or "").strip()
    if "," in texto:
        apellidos, nombres = texto.split(",", 1)
        texto = f"{nombres.strip()} {apellidos.strip()}"
    return re.sub(r"\s+", " ", texto).strip()


def extraer_nombre_tutor(campo: str) -> str:
    """Intenta obtener el nombre del tutor desde el nombre del campo."""
    texto = str(campo or "")
    patrones = [
        r"\(\s*tu[rt]or[a]?\s*:?\s*([^)]+)\)",
        r"\btu[rt]or[a]?\s*:?\s*(.+?)(?:\s*\(|$)",
    ]
    for patron in patrones:
        m = re.search(patron, texto, flags=re.IGNORECASE)
        if not m:
            continue
        nombre = m.group(1)
        nombre = re.split(r"\bdesde\b", nombre, flags=re.IGNORECASE)[0]
        nombre = re.sub(r"\s+", " ", nombre).strip(" .,-")
        if nombre:
            return nombre
    return ""


def listar_hojas(path: Path) -> list:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def buscar_hoja(wb, objetivo: str):
    objetivo_norm = normalizar(objetivo)
    for nombre in wb.sheetnames:
        if normalizar(nombre) == objetivo_norm:
            return wb[nombre]
    for nombre in wb.sheetnames:
        if objetivo_norm in normalizar(nombre):
            return wb[nombre]
    raise ValueError(f"No se encontro la hoja '{objetivo}'")


def elegir_hoja(wb, hoja=None, preferencia=None):
    if hoja:
        if hoja in wb.sheetnames:
            return wb[hoja]
        return buscar_hoja(wb, hoja)
    if preferencia:
        try:
            return buscar_hoja(wb, preferencia)
        except ValueError:
            pass
    return wb.active


def asegurar_template_docx(template=None, log=None) -> Path:
    """Devuelve un .docx listo para usar, convirtiendo el .doc con Word si hace falta."""
    log = log or print

    if template is None:
        if TEMPLATE_DOCX.exists():
            template = TEMPLATE_DOCX
        else:
            template = TEMPLATE_DOC

    template = Path(template)
    if template.suffix.lower() == ".docx":
        if not template.exists():
            raise FileNotFoundError(f"No se encontro {template}")
        return template
    if template.suffix.lower() != ".doc":
        raise ValueError("El template debe ser un archivo .doc o .docx")

    if not template.exists():
        raise FileNotFoundError(f"No se encontro {template}")

    destino = template.with_suffix(".docx")
    if destino.exists() and destino.stat().st_mtime >= template.stat().st_mtime:
        return destino

    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "Falta pywin32 para convertir el .doc. Instala con: pip install pywin32"
        ) from exc

    app = win32.Dispatch("Word.Application")
    app.Visible = False
    try:
        doc = app.Documents.Open(str(template))
        doc.SaveAs2(str(destino), FileFormat=16)
        doc.Close()
    finally:
        app.Quit()

    log(f"Template convertido a: {destino.name}")
    return destino


# ---------------------------------------------------------------------------
# MAESTRO DE ALUMNOS
# ---------------------------------------------------------------------------
def detectar_columnas_maestro(ws):
    """Detecta RUT/NOMBRE/CORREO y las columnas opcionales AREA/TIPO."""
    for fila, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
        norm = [normalizar(v) for v in row]
        if "RUT" in norm and "NOMBRE" in norm:
            cols = {"rut": norm.index("RUT"), "nombre": norm.index("NOMBRE")}
            cols["correo"] = next(
                (j for j, n in enumerate(norm) if "CORREO" in n or "MAIL" in n),
                None,
            )
            cols["area"] = next(
                (j for j, n in enumerate(norm) if n == "AREA" or n.startswith("AREA ")),
                None,
            )
            cols["tipo"] = next(
                (j for j, n in enumerate(norm) if n == "TIPO" or n.startswith("TIPO ")),
                None,
            )
            return cols, fila + 1
    return {"rut": 3, "nombre": 4, "correo": 5, "area": None, "tipo": None}, 2


def cargar_maestro(path: Path, hoja=None) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = elegir_hoja(wb, hoja, HOJA_MAESTRO)
    cols, primera_fila = detectar_columnas_maestro(ws)

    alumnos = {}
    for row in ws.iter_rows(min_row=primera_fila, values_only=True):
        def valor(indice):
            if indice is None or indice >= len(row):
                return None
            return row[indice]

        rut = valor(cols["rut"])
        nombre = valor(cols["nombre"])
        correo = valor(cols["correo"])
        area = valor(cols.get("area"))
        tipo = valor(cols.get("tipo"))

        if not nombre or str(nombre).strip() in ("", "NOMBRE"):
            continue
        alumnos[normalizar(nombre)] = {
            "nombre_original": str(nombre).strip(),
            "rut_fmt": formatear_rut(rut) if rut is not None else "",
            "correo": str(correo).strip() if correo else "",
            "area": str(area).strip() if area else "",
            "tipo": str(tipo).strip() if tipo else "",
        }
    return alumnos


def buscar_alumno(nombre_grilla: str, maestro: dict):
    clave = normalizar(nombre_grilla)
    if clave in maestro:
        return maestro[clave], "exacto"
    partes = clave.split()
    if len(partes) >= 2:
        candidatos = [k for k in maestro if partes[0] in k and partes[1] in k]
        if len(candidatos) == 1:
            return maestro[candidatos[0]], "parcial"
    return None, None


# ---------------------------------------------------------------------------
# PARSEO DE GRILLA
# ---------------------------------------------------------------------------
def es_nombre_alumno(val) -> bool:
    if not val:
        return False
    s = str(val).strip()
    if normalizar(s).lower() in IGNORAR_CELDAS:
        return False
    if re.match(r"\d{2}-\d{2}-\d{4}", s) or re.match(r"^R\d+\s*$", s):
        return False
    if len(s) < 4 or len(s) > 80:
        return False
    if re.search(r"\d", s):
        return False
    tokens = set(normalizar(s).lower().replace(".", " ").replace("-", " ").split())
    if tokens & TOKENS_NO_ALUMNO:
        return False
    return "," in s or len(s.split()) >= 2


def es_campo_clinico(val) -> bool:
    if not val:
        return False
    s = str(val).strip()
    if re.match(r"\d{2}-\d{2}-\d{4}", s) or re.match(r"^R\d+$", s.strip()):
        return False
    if s in ("Campo clinico", ""):
        return False

    norm = normalizar(s)
    if any(kw in norm.lower() for kw in PALABRAS_CLAVE_CAMPO):
        return True

    # Deteccion generica para otras carreras: cualquier etiqueta que no
    # parezca un alumno ni una nota de horario.
    base = limpiar_campo(s)
    norm_base = normalizar(base)
    if len(base) < 4 or len(base) > 140 or "," in base or re.search(r"\d", base):
        return False
    tokens = set(norm_base.lower().replace(".", " ").replace("-", " ").split())
    if tokens & TOKENS_NOTA_CAMPO:
        return False
    if norm_base.startswith("SOLO ") or norm_base.startswith("NO DISPONIBLE"):
        return False
    return True


def _tiene_fecha(valor) -> bool:
    if isinstance(valor, (date, datetime)):
        return True
    return bool(re.search(
        r"\b(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4})\b",
        str(valor or ""),
    ))


def _mapa_celdas_combinadas(ws) -> dict:
    mapa = {}
    for rango in ws.merged_cells.ranges:
        valor = ws.cell(rango.min_row, rango.min_col).value
        for fila in range(rango.min_row, rango.max_row + 1):
            for columna in range(rango.min_col, rango.max_col + 1):
                mapa[(fila, columna)] = valor
    return mapa


def parsear_grilla(path: Path, hoja=None) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = elegir_hoja(wb, hoja)
    rows = list(ws.iter_rows(values_only=True))
    combinadas = _mapa_celdas_combinadas(ws)

    def valor_celda(fila, columna):
        valor = combinadas.get((fila + 1, columna + 1), None)
        if valor is not None:
            return valor
        if fila < len(rows) and columna < len(rows[fila]):
            return rows[fila][columna]
        return None

    seccion_rows = []
    for i, row in enumerate(rows):
        vals = [str(v).strip() if v else "" for v in row]
        rot_cols = {ci: v for ci, v in enumerate(vals) if re.match(r"^R\d+$", v)}
        if len(rot_cols) >= 2:
            seccion_rows.append((i, rot_cols))

    if not seccion_rows:
        raise ValueError("No se encontraron cabeceras R1..R15 en la planilla de rotaciones")

    col_map_global = max(seccion_rows, key=lambda x: len(x[1]))[1]
    resultado = {
        rot: {
            "fecha_str": "",
            "fecha_inicio": "",
            "fecha_fin": "",
            "campos": {},
        }
        for rot in col_map_global.values()
    }

    for sec_idx, (header_i, col_map) in enumerate(seccion_rows):
        umbral = max(2, len(col_map) // 2)
        for j in range(header_i + 1, min(header_i + 4, len(rows))):
            n = sum(1 for ci in col_map if _tiene_fecha(valor_celda(j, ci)))
            if n >= umbral:
                for ci, rot in col_map.items():
                    val = valor_celda(j, ci)
                    if val is None or not rot in resultado:
                        continue
                    if resultado[rot]["fecha_str"]:
                        continue
                    rango = parsear_rango_fechas(val)
                    resultado[rot]["fecha_str"] = re.sub(
                        r"\(.*?\)", "", rango["fecha_texto_normalizado"]
                    ).strip()
                    resultado[rot]["fecha_inicio"] = rango["fecha_inicio"]
                    resultado[rot]["fecha_fin"] = rango["fecha_fin"]
                break

        fin_sec = seccion_rows[sec_idx + 1][0] if sec_idx + 1 < len(seccion_rows) else len(rows)
        campo_actual = None
        for i in range(header_i + 1, fin_sec):
            row = rows[i]
            col0 = str(row[0]).strip() if row[0] else ""
            if es_campo_clinico(col0):
                campo_actual = col0
            elif col0 and not es_nombre_alumno(col0):
                campo_actual = None
            if not campo_actual:
                continue
            for ci, rot in col_map.items():
                val = row[ci] if ci < len(row) else None
                if es_nombre_alumno(val):
                    nombre = str(val).strip()
                    if rot not in resultado:
                        resultado[rot] = {
                            "fecha_str": "",
                            "fecha_inicio": "",
                            "fecha_fin": "",
                            "campos": {},
                        }
                    resultado[rot]["campos"].setdefault(campo_actual, [])
                    if nombre not in resultado[rot]["campos"][campo_actual]:
                        resultado[rot]["campos"][campo_actual].append(nombre)

    return resultado


def ordenar_rotaciones(rots) -> list:
    return sorted(rots, key=lambda r: int(r[1:]) if r[1:].isdigit() else 99)


# ---------------------------------------------------------------------------
# MANEJO DE LA TABLA CON FILAS DINAMICAS
# ---------------------------------------------------------------------------
def set_cell_text(cell, texto: str):
    """Escribe texto en una celda conservando el formato del run original."""
    para = cell.paragraphs[0]
    if para.runs:
        para.runs[0].text = texto
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.text = texto


def clonar_fila_plantilla(tabla, fila_plantilla_idx: int):
    fila_original = tabla.rows[fila_plantilla_idx]
    nueva_tr = copy.deepcopy(fila_original._tr)
    tabla._tbl.append(nueva_tr)
    return tabla.rows[-1]


def eliminar_fila(tabla, fila_idx: int):
    tr = tabla.rows[fila_idx]._tr
    tabla._tbl.remove(tr)


def rellenar_tabla(doc: Document, alumnos_data: list, fecha_str: str, area: str):
    tabla = doc.tables[0]
    n_alumnos = len(alumnos_data)

    while len(tabla.rows) > 2:
        eliminar_fila(tabla, len(tabla.rows) - 1)

    for _ in range(n_alumnos - 1):
        clonar_fila_plantilla(tabla, 1)

    for idx, alumno in enumerate(alumnos_data):
        fila = tabla.rows[idx + 1]
        area_alumno = alumno.get("area") or area
        valores = [alumno["nombre_original"], alumno["rut_fmt"], area_alumno, fecha_str]
        for col_idx, valor in enumerate(valores):
            if col_idx < len(fila.cells):
                set_cell_text(fila.cells[col_idx], valor)


def _iter_parrafos(contenedor):
    for parrafo in contenedor.paragraphs:
        yield parrafo
    for tabla in getattr(contenedor, "tables", []):
        for fila in tabla.rows:
            for celda in fila.cells:
                yield from _iter_parrafos(celda)


def _reemplazar_posicion(runs, inicio, fin, valor):
    posicion = 0
    run_inicio = run_fin = None
    off_inicio = off_fin = 0
    for indice, run in enumerate(runs):
        largo = len(run.text)
        if run_inicio is None and posicion + largo > inicio:
            run_inicio = indice
            off_inicio = inicio - posicion
        if run_fin is None and posicion + largo >= fin:
            run_fin = indice
            off_fin = fin - posicion
            break
        posicion += largo
    if run_inicio is None or run_fin is None:
        return
    if run_inicio == run_fin:
        texto = runs[run_inicio].text
        runs[run_inicio].text = texto[:off_inicio] + valor + texto[off_fin:]
        return
    runs[run_inicio].text = runs[run_inicio].text[:off_inicio] + valor
    for indice in range(run_inicio + 1, run_fin):
        runs[indice].text = ""
    runs[run_fin].text = runs[run_fin].text[off_fin:]


def _reemplazar_en_parrafo(parrafo, variables):
    runs = parrafo.runs
    if not runs:
        return
    for variable, valor in variables.items():
        texto = "".join(run.text for run in runs)
        posiciones = []
        desde = 0
        while True:
            indice = texto.find(variable, desde)
            if indice < 0:
                break
            posiciones.append((indice, indice + len(variable)))
            desde = indice + len(variable)
        for inicio, fin in reversed(posiciones):
            _reemplazar_posicion(runs, inicio, fin, str(valor))


def reemplazar_en_doc(doc: Document, variables: dict):
    for parrafo in _iter_parrafos(doc):
        _reemplazar_en_parrafo(parrafo, variables)
    for seccion in doc.sections:
        for parte in (seccion.header, seccion.footer):
            for parrafo in _iter_parrafos(parte):
                _reemplazar_en_parrafo(parrafo, variables)


# ---------------------------------------------------------------------------
# GENERACION DE CARTA
# ---------------------------------------------------------------------------
def unir_valores(valores) -> str:
    """Une valores distintos conservando el orden (para Tipo por alumno)."""
    unicos = []
    for valor in valores:
        texto = str(valor).strip() if valor else ""
        if texto and texto not in unicos:
            unicos.append(texto)
    if not unicos:
        return ""
    if len(unicos) == 1:
        return unicos[0]
    if len(unicos) == 2:
        return f"{unicos[0]} y {unicos[1]}"
    return ", ".join(unicos[:-1]) + " y " + unicos[-1]


def generar_carta(template_path: Path, alumnos_data: list, campo: str, rotacion: str,
                  fecha_str: str, output_dir: Path, carrera: str = CARRERA,
                  tipo: str = TIPO, area: str = AREA) -> Path:
    doc = Document(str(template_path))

    hoy = datetime.today()
    fecha_envio = f"{hoy.day:02d} de {MESES[hoy.month]} de {hoy.year}"
    campo_limpio = resumir_campo(campo)

    tipos = [alumno.get("tipo") or tipo for alumno in alumnos_data]
    tipo_carta = unir_valores(tipos)

    variables = {
        "{{Fecha_Envio}}": fecha_envio,
        "{{Campo_Clinico}}": campo_limpio,
        "{{Carrera}}": carrera,
        "{{Tipo}}": tipo_carta,
    }
    reemplazar_en_doc(doc, variables)
    rellenar_tabla(doc, alumnos_data, fecha_str, area)

    campo_sf = safe_filename(campo_limpio, 40)
    ruta = output_dir / f"CARTA_{rotacion}_{campo_sf}.docx"
    doc.save(ruta)
    return ruta


# ---------------------------------------------------------------------------
# CONVERSION A PDF
# ---------------------------------------------------------------------------
def convertir_pdf(ruta_docx: Path, log=None, silenciar: bool = False):
    log = log or print
    try:
        from docx2pdf import convert
        ruta_pdf = ruta_docx.with_suffix(".pdf")
        if silenciar:
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                convert(str(ruta_docx), str(ruta_pdf))
        else:
            convert(str(ruta_docx), str(ruta_pdf))
        return ruta_pdf if ruta_pdf.exists() else None
    except ImportError:
        log("No esta instalado docx2pdf. Instala con: pip install docx2pdf")
        return None
    except Exception as e:
        log(f"No se pudo crear el PDF: {e}")
        return None


# ---------------------------------------------------------------------------
# PLANIFICACION Y PROCESO
# ---------------------------------------------------------------------------
def planificar(grilla: dict, maestro: dict, rotaciones: list):
    """Devuelve el total de cartas y el detalle por rotacion."""
    detalle = {}
    total = 0
    for rot in rotaciones:
        info = grilla.get(rot)
        if not info:
            detalle[rot] = 0
            continue
        n = 0
        for nombres in info["campos"].values():
            if any(buscar_alumno(nombre, maestro)[0] for nombre in nombres):
                n += 1
        detalle[rot] = n
        total += n
    return total, detalle


def procesar_rotacion(rotacion: str, grilla: dict, maestro: dict,
                      output_base: Path, solo_docx: bool,
                      template_path: Path, carrera: str = CARRERA,
                      tipo: str = TIPO, area: str = AREA,
                      log=None, on_carta=None, silenciar_pdf: bool = False):
    log = log or print

    if rotacion not in grilla:
        log(f"La rotacion {rotacion} no existe en la planilla.")
        return [], []

    fecha_str = grilla[rotacion]["fecha_str"]
    campos = grilla[rotacion]["campos"]

    if not campos:
        log(f"{rotacion}: sin alumnos asignados.")
        return [], []

    fecha_safe = re.sub(r"[^\w\-]", "_", fecha_str.split(" ")[0]) if fecha_str else rotacion
    rot_dir = output_base / f"{rotacion}_{fecha_safe}"
    rot_dir.mkdir(parents=True, exist_ok=True)

    log("")
    log(f"Rotacion {rotacion} - {fecha_str or 'sin fecha'}")

    generados, no_encontrados = [], []

    for campo, nombres in campos.items():
        campo_limpio = resumir_campo(campo)
        log(f"  Campo: {campo_limpio}")
        alumnos_data = []
        for nombre_grilla in nombres:
            alumno, _ = buscar_alumno(nombre_grilla, maestro)
            if not alumno:
                log(f"    No encontrado en el listado: {nombre_grilla}")
                no_encontrados.append((rotacion, campo_limpio, nombre_grilla))
                continue
            log(f"    Alumno: {alumno['nombre_original']} ({alumno['rut_fmt']})")
            alumnos_data.append(alumno)

        if not alumnos_data:
            log("    Sin alumnos validos para este campo. Se omite la carta.")
            continue

        ruta_docx = generar_carta(
            template_path, alumnos_data, campo, rotacion, fecha_str, rot_dir,
            carrera, tipo, area,
        )
        ruta_final = ruta_docx

        if not solo_docx:
            ruta_pdf = convertir_pdf(ruta_docx, log=log, silenciar=silenciar_pdf)
            if ruta_pdf and ruta_pdf.exists():
                ruta_docx.unlink()
                ruta_final = ruta_pdf
            else:
                log("    Se conserva el archivo DOCX porque el PDF no se pudo generar.")
        log(f"    Carta generada: {ruta_final.name}")

        generados.append({
            "rotacion": rotacion,
            "fecha": fecha_str,
            "campo": campo_limpio,
            "campo_original": campo,
            "alumnos": ", ".join(a["nombre_original"] for a in alumnos_data),
            "correos": "; ".join(a["correo"] for a in alumnos_data if a["correo"]),
            "archivo": str(ruta_final),
        })
        if on_carta:
            on_carta()

    return generados, no_encontrados


# ---------------------------------------------------------------------------
# MAIN CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Genera cartas PDF de practicas/preclinicas IPCHILE"
    )
    grupo = parser.add_mutually_exclusive_group(required=True)
    grupo.add_argument("--rotacion", nargs="+", metavar="R#")
    grupo.add_argument("--all", action="store_true")
    parser.add_argument("--output", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--solo-docx", action="store_true", help="No convertir a PDF")
    parser.add_argument("--lista", type=Path, default=MAESTRO_XLSX)
    parser.add_argument("--grilla", type=Path, default=GRILLA_XLSX)
    parser.add_argument("--hoja-lista", default=None)
    parser.add_argument("--hoja-grilla", default=None)
    parser.add_argument("--template", type=Path, default=None)
    parser.add_argument("--carrera", default=CARRERA)
    parser.add_argument("--tipo", default=TIPO)
    parser.add_argument("--area", default=AREA)
    parser.add_argument("--planilla", action="store_true",
                        help="Generar tambien la planilla de correos en Excel")
    parser.add_argument("--planilla-salida", type=Path, default=None)
    parser.add_argument("--adjuntos-comunes", nargs="*", default=None)
    parser.add_argument("--correo-template", type=Path, default=None)
    args = parser.parse_args()

    faltantes = [f for f in [args.lista, args.grilla] if not f.exists()]
    if faltantes:
        for f in faltantes:
            print(f"No encontrado: {f.name}")
        print("Revisa las rutas o usa la interfaz para seleccionar los archivos.")
        sys.exit(1)

    try:
        template_path = asegurar_template_docx(args.template)
    except Exception as e:
        print(f"Template: {e}")
        sys.exit(1)

    print("Cargando listado de alumnos...")
    maestro = cargar_maestro(args.lista, args.hoja_lista)
    print(f"{len(maestro)} alumnos cargados.")

    print("Leyendo grilla de rotaciones...")
    grilla = parsear_grilla(args.grilla, args.hoja_grilla)
    rots = ordenar_rotaciones(grilla.keys())
    print(f"Rotaciones: {', '.join(rots)}")

    rots_procesar = rots if args.all else [r.upper() for r in args.rotacion]
    args.output.mkdir(parents=True, exist_ok=True)

    todos_gen, todos_nenc = [], []
    for rot in rots_procesar:
        gen, nenc = procesar_rotacion(
            rot, grilla, maestro, args.output, args.solo_docx,
            template_path, args.carrera, args.tipo, args.area,
        )
        todos_gen.extend(gen)
        todos_nenc.extend(nenc)

    if args.planilla:
        try:
            import planilla_correos as pc
            destino = args.planilla_salida or (args.output / "Planilla_Correos.xlsx")
            correo_template = args.correo_template
            if correo_template is None:
                candidato = BASE_DIR / "Correo_Template.docx"
                correo_template = candidato if candidato.exists() else None
            resultado = pc.exportar_planilla(
                grilla, maestro, rots_procesar, args.output, destino,
                args.carrera, args.adjuntos_comunes, todos_gen, correo_template,
            )
            print(f"Planilla de correos: {resultado['archivo']}")
            if resultado["campos_faltantes_template"]:
                print("Campos del template sin columna: "
                      + ", ".join(resultado["campos_faltantes_template"]))
        except Exception as e:
            print(f"No se pudo generar la planilla de correos: {e}")

    ext = "docx" if args.solo_docx else "PDF"
    print("")
    print("=" * 60)
    print("RESUMEN FINAL")
    print("=" * 60)
    print(f"Cartas generadas ({ext}): {len(todos_gen)}")
    print(f"No encontrados: {len(todos_nenc)}")
    if todos_nenc:
        print("")
        print("Sin match en el listado:")
        for rot, campo, nombre in todos_nenc:
            print(f"  [{rot}] {campo}: {nombre}")
    print(f"Carpeta de salida: {args.output.resolve()}")
    print("")


if __name__ == "__main__":
    main()
