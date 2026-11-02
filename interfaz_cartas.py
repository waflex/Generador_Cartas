#!/usr/bin/env python3
"""
interfaz_cartas.py
Interfaz grafica para generar cartas de practica a partir de:
  - Listado de alumnos (.xlsx)
  - Grilla de rotaciones (.xlsx)
  - Template de carta (.doc o .docx)

Ejecutar:
    python interfaz_cartas.py
    pythonw interfaz_cartas.py   (sin ventana de terminal)
"""

import os
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import config as app_config
import catalogo_tutores as ct
import generar_cartas as g
import planilla_correos as pc

COLOR_FONDO = "#F4F6FA"
COLOR_TARJETA = "#FFFFFF"
COLOR_BORDE = "#E2E8F0"
COLOR_TEXTO = "#0F172A"
COLOR_TEXTO_SUAVE = "#64748B"
COLOR_ACENTO = "#2563EB"
COLOR_ACENTO_OSCURO = "#1D4ED8"
COLOR_ACENTO_CLARO = "#DBEAFE"
COLOR_SIDEBAR = "#0F172A"
COLOR_SIDEBAR_TARJETA = "#1E293B"
FUENTE = "Segoe UI"


class Aplicacion:
    def __init__(self, root):
        self.root = root
        self.root.title("Cruce de Correos | Generador de Cartas")
        ancho = min(1180, max(1100, self.root.winfo_screenwidth() - 80))
        alto = min(820, max(620, self.root.winfo_screenheight() - 100))
        self.root.geometry(f"{ancho}x{alto}")
        self.root.minsize(1100, 620)
        self.root.configure(bg=COLOR_FONDO)

        self.cola = queue.Queue()
        self.ejecutando = False
        self.grilla = {}
        self.grilla_fuente = None
        self.codigos_rot = []
        self.ultimos_generados = []
        self.rutas_comunes = []
        self.pasos = []
        self.ventana_config = None
        self.ventana_dic = None
        self.arbol_dic = None
        self.var_busqueda_dic = None
        self.lbl_dic_conteo = None
        app_config.asegurar_config()

        base = g.BASE_DIR
        self.var_lista = tk.StringVar()
        self.var_grilla = tk.StringVar()
        self.var_plantilla = tk.StringVar()
        self.var_correo_template = tk.StringVar()
        self.var_correo_recordatorio = tk.StringVar()
        self.var_catalogo = tk.StringVar()
        self.var_comunes = tk.StringVar()
        self.var_salida = tk.StringVar(value=str(base / "cartas_output"))
        self.var_hoja_lista = tk.StringVar()
        self.var_hoja_grilla = tk.StringVar()
        self.var_hoja_catalogo = tk.StringVar()
        self.var_carrera = tk.StringVar(value=g.CARRERA)
        self.var_tipo = tk.StringVar(value=g.TIPO)
        self.var_area = tk.StringVar(value=g.AREA)
        self.var_hora_ingreso = tk.StringVar(
            value=g.CONFIG.get("recordatorio", {}).get("hora_ingreso", "09:00")
        )
        self.var_pdf = tk.BooleanVar(value=True)
        self.var_modo_borrador = tk.StringVar(
            value="Por grupo: presentacion + recordatorio"
        )
        self.var_cco = tk.BooleanVar(value=False)
        self.var_cc_tutor = tk.BooleanVar(value=True)
        self.var_mostrar = tk.BooleanVar(value=True)
        self.var_estado = tk.StringVar(value="Listo para comenzar")
        self.var_resumen = tk.StringVar(value="")

        if g.MAESTRO_XLSX.exists():
            self.var_lista.set(str(g.MAESTRO_XLSX))
        if g.GRILLA_XLSX.exists():
            self.var_grilla.set(str(g.GRILLA_XLSX))
        for candidato in (g.TEMPLATE_DOCX, g.TEMPLATE_DOC):
            if candidato.exists():
                self.var_plantilla.set(str(candidato))
                break
        for candidato in (
            app_config.ruta_recurso("Correo_Template.docx"),
            app_config.ruta_recurso("Correo_template.docx"),
        ):
            if candidato.exists():
                self.var_correo_template.set(str(candidato))
                break
        recordatorio = app_config.ruta_recurso(
            "Estimada Estudiante recordatorio.docx"
        )
        if recordatorio.exists():
            self.var_correo_recordatorio.set(str(recordatorio))

        for var in (
            self.var_lista, self.var_hoja_lista,
            self.var_grilla, self.var_hoja_grilla,
            self.var_plantilla, self.var_correo_template,
            self.var_correo_recordatorio, self.var_catalogo,
            self.var_hoja_catalogo, self.var_comunes, self.var_salida,
        ):
            var.trace_add("write", self._resumen_cambiado)

        self._configurar_estilo()
        self._construir()
        self._crear_ventana_config()
        self._cargar_hojas_iniciales()
        self.root.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.root.after(100, self._procesar_cola)

    # ------------------------------------------------------------------
    # ESTILO Y ESTRUCTURA
    # ------------------------------------------------------------------
    def _configurar_estilo(self):
        estilo = ttk.Style(self.root)
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass

        estilo.configure(
            ".", font=(FUENTE, 10), background=COLOR_FONDO,
            foreground=COLOR_TEXTO,
        )
        estilo.configure("TFrame", background=COLOR_FONDO)
        estilo.configure("TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO)
        estilo.configure(
            "Card.TLabel", background=COLOR_TARJETA, foreground=COLOR_TEXTO
        )
        estilo.configure(
            "CardMuted.TLabel", background=COLOR_TARJETA,
            foreground=COLOR_TEXTO_SUAVE, font=(FUENTE, 9),
        )
        estilo.configure(
            "Count.TLabel", background=COLOR_TARJETA,
            foreground=COLOR_ACENTO, font=(FUENTE, 10, "bold"),
        )
        estilo.configure(
            "Title.TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO,
            font=(FUENTE, 20, "bold"),
        )
        estilo.configure(
            "Subtitle.TLabel", background=COLOR_FONDO,
            foreground=COLOR_TEXTO_SUAVE, font=(FUENTE, 10),
        )

        estilo.configure(
            "Accent.TButton", background=COLOR_ACENTO, foreground="#FFFFFF",
            borderwidth=0, focusthickness=0, padding=(18, 10),
            font=(FUENTE, 10, "bold"),
        )
        estilo.map(
            "Accent.TButton",
            background=[
                ("pressed", COLOR_ACENTO_OSCURO),
                ("active", COLOR_ACENTO_OSCURO),
                ("disabled", "#93C5FD"),
            ],
            foreground=[("disabled", "#F8FAFC")],
        )
        estilo.configure(
            "Secondary.TButton", background=COLOR_ACENTO_CLARO,
            foreground=COLOR_ACENTO, borderwidth=0, focusthickness=0,
            padding=(14, 9), font=(FUENTE, 10, "bold"),
        )
        estilo.map(
            "Secondary.TButton",
            background=[
                ("pressed", "#93C5FD"),
                ("active", "#BFDBFE"),
                ("disabled", "#EFF6FF"),
            ],
            foreground=[("disabled", "#93C5FD")],
        )
        estilo.configure(
            "Ghost.TButton", background=COLOR_FONDO,
            foreground=COLOR_TEXTO_SUAVE, borderwidth=0, focusthickness=0,
            padding=(10, 8), font=(FUENTE, 10),
        )
        estilo.map(
            "Ghost.TButton",
            background=[("active", "#E2E8F0")],
            foreground=[("active", COLOR_TEXTO)],
        )
        estilo.configure(
            "Small.TButton", background="#F1F5F9", foreground=COLOR_TEXTO,
            borderwidth=0, focusthickness=0, padding=(10, 5),
            font=(FUENTE, 9, "bold"),
        )
        estilo.map(
            "Small.TButton", background=[("active", "#E2E8F0")]
        )

        estilo.configure(
            "TEntry", fieldbackground=COLOR_TARJETA, foreground=COLOR_TEXTO,
            bordercolor=COLOR_BORDE, lightcolor=COLOR_BORDE,
            darkcolor=COLOR_BORDE, padding=(8, 6),
        )
        estilo.map(
            "TEntry",
            bordercolor=[("focus", COLOR_ACENTO)],
            lightcolor=[("focus", COLOR_ACENTO)],
            darkcolor=[("focus", COLOR_ACENTO)],
        )
        estilo.configure(
            "Path.TEntry", font=(FUENTE, 9), padding=(6, 5)
        )
        estilo.configure(
            "TCombobox", fieldbackground=COLOR_TARJETA,
            background=COLOR_TARJETA, foreground=COLOR_TEXTO,
            arrowcolor=COLOR_TEXTO, bordercolor=COLOR_BORDE,
            padding=(8, 5), font=(FUENTE, 9),
        )
        estilo.map(
            "TCombobox",
            fieldbackground=[("readonly", COLOR_TARJETA)],
            selectbackground=[("readonly", COLOR_TARJETA)],
            selectforeground=[("readonly", COLOR_TEXTO)],
        )
        estilo.configure(
            "TCheckbutton", background=COLOR_TARJETA,
            foreground=COLOR_TEXTO, font=(FUENTE, 10),
        )
        estilo.map(
            "TCheckbutton",
            background=[("active", COLOR_TARJETA)],
            foreground=[("active", COLOR_TEXTO)],
        )
        estilo.configure("Vertical.TScrollbar", background="#CBD5E1",
                         troughcolor=COLOR_TARJETA, bordercolor=COLOR_TARJETA,
                         arrowcolor=COLOR_TEXTO_SUAVE)
        estilo.configure(
            "Log.Vertical.TScrollbar", background="#334155",
            troughcolor="#0F172A", bordercolor="#0F172A",
            arrowcolor="#94A3B8",
        )
        estilo.configure(
            "TNotebook", background=COLOR_FONDO, borderwidth=0,
            tabmargins=(0, 0, 0, 0),
        )
        estilo.configure(
            "TNotebook.Tab", background=COLOR_FONDO,
            foreground=COLOR_TEXTO_SUAVE, padding=(20, 10),
            font=(FUENTE, 10, "bold"), borderwidth=0,
        )
        estilo.map(
            "TNotebook.Tab",
            background=[("selected", COLOR_TARJETA)],
            foreground=[("selected", COLOR_ACENTO)],
        )
        estilo.configure(
            "Horizontal.TProgressbar", troughcolor="#E2E8F0",
            background=COLOR_ACENTO, borderwidth=0, thickness=10,
        )

    def _card(
        self, padre, titulo, subtitulo=None, fila=0, columna=0,
        padx=0, pady=0,
    ):
        tarjeta = tk.Frame(
            padre, bg=COLOR_TARJETA, highlightbackground=COLOR_BORDE,
            highlightthickness=1,
        )
        tarjeta.grid(
            row=fila, column=columna, sticky="nsew", padx=padx, pady=pady
        )
        cabecera = tk.Frame(tarjeta, bg=COLOR_TARJETA)
        cabecera.pack(fill="x", padx=20, pady=(16, 0))
        tk.Label(
            cabecera, text=titulo, bg=COLOR_TARJETA, fg=COLOR_TEXTO,
            font=(FUENTE, 11, "bold"),
        ).pack(anchor="w")
        if subtitulo:
            tk.Label(
                cabecera, text=subtitulo, bg=COLOR_TARJETA,
                fg=COLOR_TEXTO_SUAVE, font=(FUENTE, 9),
            ).pack(anchor="w", pady=(3, 0))
        cuerpo = tk.Frame(tarjeta, bg=COLOR_TARJETA)
        cuerpo.pack(fill="both", expand=True, padx=20, pady=(12, 18))
        return cuerpo

    def _tip(self, padre, texto, wraplength=220):
        caja = tk.Frame(
            padre, bg="#EFF6FF", highlightbackground="#BFDBFE",
            highlightthickness=1,
        )
        tk.Frame(caja, bg=COLOR_ACENTO, width=3).pack(side="left", fill="y")
        tk.Label(
            caja, text=texto, bg="#EFF6FF", fg="#1E3A8A",
            font=(FUENTE, 9), wraplength=wraplength, justify="left",
            anchor="w", padx=10, pady=8,
        ).pack(side="left", fill="x", expand=True)
        return caja

    def _crear_paso(self, padre, numero, texto):
        fila = tk.Frame(padre, bg=COLOR_SIDEBAR)
        fila.pack(fill="x", pady=4)
        etiqueta_numero = tk.Label(
            fila, text=numero, bg=COLOR_SIDEBAR_TARJETA, fg="#93C5FD",
            font=(FUENTE, 8, "bold"), width=3, pady=4,
        )
        etiqueta_numero.pack(side="left")
        etiqueta_texto = tk.Label(
            fila, text=texto, bg=COLOR_SIDEBAR, fg="#94A3B8",
            font=(FUENTE, 10),
        )
        etiqueta_texto.pack(side="left", padx=10)
        self.pasos.append((etiqueta_numero, etiqueta_texto))

    def _marcar_pasos(self, activos):
        for indice, (numero, texto) in enumerate(self.pasos):
            if indice in activos:
                numero.configure(bg=COLOR_ACENTO, fg="#FFFFFF")
                texto.configure(fg="#FFFFFF")
            else:
                numero.configure(bg=COLOR_SIDEBAR_TARJETA, fg="#93C5FD")
                texto.configure(fg="#94A3B8")

    def _cambiar_paso(self, _evento=None):
        if (
            self.ventana_config is not None
            and self.ventana_config.winfo_exists()
            and self.ventana_config.winfo_viewable()
        ):
            self._marcar_pasos({0})
            return
        try:
            actual = self.notebook.index("current")
        except tk.TclError:
            return
        activos = {0: {1}, 1: {2}}.get(actual, set())
        self._marcar_pasos(activos)

    def _construir(self):
        self.root.columnconfigure(0, weight=0)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        self._construir_sidebar()

        principal = tk.Frame(self.root, bg=COLOR_FONDO)
        principal.grid(row=0, column=1, sticky="nsew")
        principal.columnconfigure(0, weight=1)
        principal.rowconfigure(1, weight=1)

        self._construir_cabecera(principal)
        self._construir_notebook(principal)
        tk.Frame(principal, bg=COLOR_BORDE, height=1).grid(
            row=2, column=0, sticky="ew", padx=28
        )
        self._construir_botones(principal)

    def _construir_sidebar(self):
        sidebar = tk.Frame(self.root, bg=COLOR_SIDEBAR, width=262)
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)
        sidebar.columnconfigure(0, weight=1)
        sidebar.rowconfigure(4, weight=1)

        marca = tk.Frame(sidebar, bg=COLOR_SIDEBAR)
        marca.grid(row=0, column=0, sticky="ew", padx=24, pady=(28, 22))
        tk.Label(
            marca, text="IPCHILE", bg=COLOR_SIDEBAR, fg="#60A5FA",
            font=(FUENTE, 9, "bold"),
        ).pack(anchor="w")
        tk.Label(
            marca, text="Cruce de Correos", bg=COLOR_SIDEBAR, fg="#FFFFFF",
            font=(FUENTE, 18, "bold"),
        ).pack(anchor="w", pady=(6, 0))
        tk.Label(
            marca, text="Generador de cartas", bg=COLOR_SIDEBAR,
            fg="#94A3B8", font=(FUENTE, 9),
        ).pack(anchor="w", pady=(4, 0))

        tk.Frame(sidebar, bg="#1E293B", height=1).grid(
            row=1, column=0, sticky="ew", padx=24
        )
        tk.Label(
            sidebar, text="FLUJO DE TRABAJO", bg=COLOR_SIDEBAR,
            fg="#64748B", font=(FUENTE, 8, "bold"),
        ).grid(row=2, column=0, sticky="w", padx=24, pady=(22, 10))

        pasos = tk.Frame(sidebar, bg=COLOR_SIDEBAR)
        pasos.grid(row=3, column=0, sticky="ew", padx=24)
        for numero, texto in (
            ("01", "Configurar"), ("02", "Rotaciones"),
            ("03", "Proceso"),
        ):
            self._crear_paso(pasos, numero, texto)

        estado = tk.Frame(sidebar, bg=COLOR_SIDEBAR_TARJETA)
        estado.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 16))
        tk.Label(
            estado, text="ESTADO", bg=COLOR_SIDEBAR_TARJETA,
            fg="#94A3B8", font=(FUENTE, 8, "bold"),
        ).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(
            estado, textvariable=self.var_estado, bg=COLOR_SIDEBAR_TARJETA,
            fg="#FFFFFF", font=(FUENTE, 9), wraplength=198, justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 12))

        tk.Label(
            sidebar, text="Prototipo v1.0", bg=COLOR_SIDEBAR, fg="#475569",
            font=(FUENTE, 8),
        ).grid(row=6, column=0, sticky="w", padx=24, pady=(0, 18))

    def _construir_cabecera(self, padre):
        cabecera = tk.Frame(padre, bg=COLOR_FONDO)
        cabecera.grid(row=0, column=0, sticky="ew", padx=28, pady=(22, 12))
        cabecera.columnconfigure(0, weight=1)
        tk.Label(
            cabecera, text="Panel de generacion", bg=COLOR_FONDO,
            fg=COLOR_TEXTO, font=(FUENTE, 20, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            cabecera,
            text="Selecciona las rotaciones y genera las cartas.",
            bg=COLOR_FONDO, fg=COLOR_TEXTO_SUAVE, font=(FUENTE, 10),
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        acciones = tk.Frame(cabecera, bg=COLOR_FONDO)
        acciones.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Button(
            acciones, text="Diccionario", style="Small.TButton",
            command=self._abrir_diccionario,
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            acciones, text="Configuracion", style="Secondary.TButton",
            command=self._abrir_configuracion,
        ).pack(side="left")

        resumen = tk.Frame(
            cabecera, bg=COLOR_TARJETA, highlightbackground=COLOR_BORDE,
            highlightthickness=1,
        )
        resumen.grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(14, 0)
        )
        self.lbl_resumen = tk.Label(
            resumen, textvariable=self.var_resumen, bg=COLOR_TARJETA,
            fg=COLOR_TEXTO_SUAVE, font=(FUENTE, 9), anchor="w",
            justify="left", wraplength=790, padx=12, pady=8,
        )
        self.lbl_resumen.pack(side="left", fill="x", expand=True)
        ttk.Button(
            resumen, text="Editar", style="Small.TButton",
            command=self._abrir_configuracion,
        ).pack(side="right", padx=8, pady=6)

    def _construir_notebook(self, padre):
        self.notebook = ttk.Notebook(padre)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=28)
        self.tab_rotaciones = tk.Frame(
            self.notebook, bg=COLOR_FONDO, padx=2, pady=14
        )
        self.tab_proceso = tk.Frame(
            self.notebook, bg=COLOR_FONDO, padx=2, pady=14
        )
        self.notebook.add(self.tab_rotaciones, text="Rotaciones")
        self.notebook.add(self.tab_proceso, text="Proceso")

        self.tab_rotaciones.columnconfigure(0, weight=4)
        self.tab_rotaciones.columnconfigure(1, weight=2)
        self.tab_rotaciones.rowconfigure(0, weight=1)
        self.tab_proceso.columnconfigure(0, weight=1)
        self.tab_proceso.rowconfigure(0, weight=1)

        self._construir_rotaciones(self.tab_rotaciones)
        self._construir_opciones_outlook(self.tab_rotaciones)
        self._construir_progreso(self.tab_proceso)

        self.notebook.bind("<<NotebookTabChanged>>", self._cambiar_paso)
        self._cambiar_paso()

    def _resumen_cambiado(self, *_args):
        self._actualizar_resumen()

    def _actualizar_resumen(self):
        if not hasattr(self, "lbl_resumen"):
            return
        partes = []
        for etiqueta, var_ruta, var_hoja in (
            ("Listado", self.var_lista, self.var_hoja_lista),
            ("Planilla", self.var_grilla, self.var_hoja_grilla),
            ("Catalogo", self.var_catalogo, self.var_hoja_catalogo),
        ):
            ruta = var_ruta.get().strip()
            if not ruta:
                partes.append(f"{etiqueta}: sin definir")
                continue
            hoja = var_hoja.get().strip()
            nombre = Path(ruta).name
            partes.append(
                f"{etiqueta}: {nombre}" + (f" [{hoja}]" if hoja else "")
            )
        plantilla = self.var_plantilla.get().strip()
        partes.append(
            "Cartas: "
            + (Path(plantilla).name if plantilla else "sin definir")
        )
        salida = self.var_salida.get().strip()
        partes.append("Salida: " + (salida if salida else "sin definir"))
        if self.rutas_comunes:
            partes.append(f"Adjuntos: {len(self.rutas_comunes)}")
        self.var_resumen.set("   |   ".join(partes))

    def _crear_ventana_config(self):
        if (
            self.ventana_config is not None
            and self.ventana_config.winfo_exists()
        ):
            return
        ventana = tk.Toplevel(self.root)
        ventana.title("Configuracion general")
        ventana.geometry("1020x640")
        ventana.minsize(900, 560)
        ventana.configure(bg=COLOR_FONDO)
        ventana.transient(self.root)
        ventana.protocol("WM_DELETE_WINDOW", self._cerrar_config)
        ventana.withdraw()
        self.ventana_config = ventana

        cabecera = tk.Frame(ventana, bg=COLOR_FONDO)
        cabecera.pack(fill="x", padx=20, pady=(18, 10))
        tk.Label(
            cabecera, text="Configuracion general", bg=COLOR_FONDO,
            fg=COLOR_TEXTO, font=(FUENTE, 16, "bold"),
        ).pack(side="left")
        ttk.Button(
            cabecera, text="Cerrar", style="Small.TButton",
            command=self._cerrar_config,
        ).pack(side="right")

        contenedor = tk.Frame(ventana, bg=COLOR_FONDO)
        contenedor.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        contenedor.columnconfigure(0, weight=3)
        contenedor.columnconfigure(1, weight=2)
        contenedor.rowconfigure(0, weight=1)

        self._construir_archivos(contenedor)
        self._construir_datos(contenedor)
        self._actualizar_resumen()

    def _abrir_configuracion(self):
        if (
            self.ventana_config is None
            or not self.ventana_config.winfo_exists()
        ):
            self._crear_ventana_config()
        self._marcar_pasos({0})
        self.ventana_config.deiconify()
        self.ventana_config.lift()
        self.ventana_config.focus_force()

    def _cerrar_config(self):
        if (
            self.ventana_config is not None
            and self.ventana_config.winfo_exists()
        ):
            self.ventana_config.withdraw()
        self._actualizar_resumen()
        self._cambiar_paso()

    def _abrir_diccionario(self):
        if self.ventana_dic is not None and self.ventana_dic.winfo_exists():
            self.ventana_dic.deiconify()
            self.ventana_dic.lift()
            self._filtrar_diccionario()
            return

        ventana = tk.Toplevel(self.root)
        ventana.title("Diccionario de campos clinicos")
        ventana.geometry("780x560")
        ventana.minsize(620, 420)
        ventana.configure(bg=COLOR_FONDO)
        ventana.transient(self.root)
        ventana.protocol("WM_DELETE_WINDOW", self._cerrar_diccionario)
        self.ventana_dic = ventana

        cabecera = tk.Frame(ventana, bg=COLOR_FONDO)
        cabecera.pack(fill="x", padx=18, pady=(16, 8))
        tk.Label(
            cabecera, text="Diccionario de campos clinicos",
            bg=COLOR_FONDO, fg=COLOR_TEXTO, font=(FUENTE, 15, "bold"),
        ).pack(anchor="w")
        tk.Label(
            cabecera,
            text="Nombres completos y abreviaturas usadas en cartas y correos.",
            bg=COLOR_FONDO, fg=COLOR_TEXTO_SUAVE, font=(FUENTE, 9),
        ).pack(anchor="w", pady=(3, 0))

        filtro = tk.Frame(ventana, bg=COLOR_FONDO)
        filtro.pack(fill="x", padx=18, pady=(4, 10))
        tk.Label(
            filtro, text="Buscar", bg=COLOR_FONDO, fg=COLOR_TEXTO_SUAVE,
            font=(FUENTE, 9),
        ).pack(side="left", padx=(0, 8))
        self.var_busqueda_dic = tk.StringVar()
        entrada = ttk.Entry(
            filtro, textvariable=self.var_busqueda_dic,
        )
        entrada.pack(side="left", fill="x", expand=True)
        entrada.bind(
            "<KeyRelease>", lambda _e: self._filtrar_diccionario()
        )
        entrada.focus_set()
        self.lbl_dic_conteo = tk.Label(
            filtro, text="", bg=COLOR_FONDO, fg=COLOR_TEXTO_SUAVE,
            font=(FUENTE, 9),
        )
        self.lbl_dic_conteo.pack(side="right", padx=(10, 0))

        marco = tk.Frame(ventana, bg=COLOR_FONDO)
        marco.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(0, weight=1)

        self.arbol_dic = ttk.Treeview(
            marco, columns=("Campo", "Abreviatura"), show="headings",
            selectmode="browse",
        )
        self.arbol_dic.heading(
            "Campo", text="Campo clinico", anchor="w"
        )
        self.arbol_dic.heading(
            "Abreviatura", text="Abreviatura", anchor="w"
        )
        self.arbol_dic.column("Campo", width=560, anchor="w")
        self.arbol_dic.column("Abreviatura", width=130, anchor="w")
        self.arbol_dic.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(
            marco, orient="vertical", command=self.arbol_dic.yview
        )
        barra.grid(row=0, column=1, sticky="ns")
        self.arbol_dic.configure(yscrollcommand=barra.set)

        pie = tk.Frame(ventana, bg=COLOR_FONDO)
        pie.pack(fill="x", padx=18, pady=(0, 14))
        ttk.Button(
            pie, text="Recargar config.json", style="Small.TButton",
            command=self._recargar_diccionario,
        ).pack(side="left")
        tk.Label(
            pie, text="Se edita en config.json", bg=COLOR_FONDO,
            fg=COLOR_TEXTO_SUAVE, font=(FUENTE, 9),
        ).pack(side="left", padx=(10, 0))
        ttk.Button(
            pie, text="Cerrar", style="Small.TButton",
            command=self._cerrar_diccionario,
        ).pack(side="right")

        self._filtrar_diccionario()

    def _cerrar_diccionario(self):
        if self.ventana_dic is not None:
            self.ventana_dic.destroy()
            self.ventana_dic = None
            self.arbol_dic = None

    def _filtrar_diccionario(self):
        if self.arbol_dic is None or not self.arbol_dic.winfo_exists():
            return
        consulta = g.normalizar(
            self.var_busqueda_dic.get() if self.var_busqueda_dic else ""
        )
        for item in self.arbol_dic.get_children():
            self.arbol_dic.delete(item)
        filas = sorted(
            g.CAMPOS_CLINICOS.items(),
            key=lambda item: g.normalizar(item[1]),
        )
        mostrados = 0
        for campo, abreviatura in filas:
            if consulta and (
                consulta not in g.normalizar(campo)
                and consulta not in g.normalizar(abreviatura)
            ):
                continue
            self.arbol_dic.insert(
                "", "end", values=(campo, abreviatura)
            )
            mostrados += 1
        if self.lbl_dic_conteo is not None:
            self.lbl_dic_conteo.config(text=f"{mostrados} campos")

    def _recargar_diccionario(self):
        g.recargar_config()
        self._filtrar_diccionario()
        self._actualizar_resumen()

    def _construir_archivos(self, padre):
        marco = self._card(
            padre,
            "Archivos de entrada",
            "Selecciona el listado, la planilla y las plantillas.",
            fila=0, columna=0, padx=(0, 7),
        )
        marco.columnconfigure(1, weight=1)

        self._fila_archivo(
            marco, 0, "Listado de alumnos",
            self.var_lista, self.var_hoja_lista,
            lambda: self._elegir_excel(
                self.var_lista, self.combo_hoja_lista, g.HOJA_MAESTRO, True
            ),
        )
        self.combo_hoja_lista = self._combo_hoja
        self._fila_archivo(
            marco, 1, "Planilla de rotaciones",
            self.var_grilla, self.var_hoja_grilla,
            lambda: self._elegir_excel(
                self.var_grilla, self.combo_hoja_grilla, None, True
            ),
        )
        self.combo_hoja_grilla = self._combo_hoja
        self.combo_hoja_grilla.bind(
            "<<ComboboxSelected>>", lambda _e: self._recargar_rotaciones()
        )

        self._fila_archivo(
            marco, 2, "Catalogo de tutores",
            self.var_catalogo, self.var_hoja_catalogo,
            lambda: self._elegir_excel(
                self.var_catalogo, self.combo_hoja_catalogo, None, False
            ),
        )
        self.combo_hoja_catalogo = self._combo_hoja

        self._fila_simple(
            marco, 3, "Plantilla de carta", self.var_plantilla,
            self._elegir_template,
        )
        self._fila_simple(
            marco, 4, "Carpeta de salida", self.var_salida,
            self._elegir_salida, "Elegir",
        )
        self._fila_simple(
            marco, 5, "Plantilla de correo",
            self.var_correo_template, self._elegir_correo_template,
        )
        self._fila_simple(
            marco, 6, "Plantilla de recordatorio",
            self.var_correo_recordatorio, self._elegir_correo_recordatorio,
        )

        ttk.Label(
            marco, text="Adjuntos comunes", style="Card.TLabel"
        ).grid(row=7, column=0, sticky="w", pady=5)
        ttk.Entry(
            marco, textvariable=self.var_comunes, state="readonly",
            width=26, style="Path.TEntry",
        ).grid(row=7, column=1, sticky="ew", padx=(0, 8), pady=5)
        acciones = tk.Frame(marco, bg=COLOR_TARJETA)
        acciones.grid(row=7, column=2, columnspan=2, sticky="w", pady=5)
        ttk.Button(
            acciones, text="Elegir archivos", style="Small.TButton",
            command=self._elegir_comunes,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            acciones, text="Quitar", style="Small.TButton",
            command=self._quitar_comunes,
        ).pack(side="left")

        self._tip(
            marco,
            "El catalogo de tutores es opcional. Los nombres cortos se "
            "configuran en config.json.",
            wraplength=500,
        ).grid(row=8, column=0, columnspan=4, sticky="ew", pady=(16, 0))

    def _fila_archivo(self, marco, fila, etiqueta, var_ruta, var_hoja, comando):
        ttk.Label(
            marco, text=etiqueta, style="Card.TLabel"
        ).grid(row=fila, column=0, sticky="w", pady=5)
        ttk.Entry(
            marco, textvariable=var_ruta, width=26, style="Path.TEntry"
        ).grid(row=fila, column=1, sticky="ew", padx=(0, 8), pady=5)
        ttk.Button(
            marco, text="Examinar", style="Small.TButton", command=comando
        ).grid(row=fila, column=2, sticky="w", pady=5)
        self._combo_hoja = ttk.Combobox(
            marco, textvariable=var_hoja, state="readonly", width=16
        )
        self._combo_hoja.grid(row=fila, column=3, padx=(8, 0), pady=5)

    def _fila_simple(
        self, marco, fila, etiqueta, var_ruta, comando, texto_boton="Examinar"
    ):
        ttk.Label(
            marco, text=etiqueta, style="Card.TLabel"
        ).grid(row=fila, column=0, sticky="w", pady=5)
        ttk.Entry(
            marco, textvariable=var_ruta, width=26, style="Path.TEntry"
        ).grid(row=fila, column=1, sticky="ew", padx=(0, 8), pady=5)
        ttk.Button(
            marco, text=texto_boton, style="Small.TButton", command=comando
        ).grid(row=fila, column=2, sticky="w", pady=5)

    def _construir_datos(self, padre):
        marco = self._card(
            padre,
            "Parametros de la carrera",
            "AREA y TIPO se leen del listado.",
            fila=0, columna=1, padx=(7, 0),
        )
        marco.columnconfigure(1, weight=1)

        ttk.Label(
            marco, text="Carrera", style="Card.TLabel"
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(marco, textvariable=self.var_carrera).grid(
            row=0, column=1, sticky="ew", pady=(0, 6)
        )
        ttk.Label(
            marco, text="Area", style="Card.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(marco, textvariable=self.var_area).grid(
            row=1, column=1, sticky="ew", pady=(0, 6)
        )
        ttk.Label(
            marco, text="Tipo", style="Card.TLabel"
        ).grid(row=2, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(marco, textvariable=self.var_tipo).grid(
            row=2, column=1, sticky="ew", pady=(0, 6)
        )

        ttk.Checkbutton(
            marco,
            text="Generar PDF con Word",
            variable=self.var_pdf,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        self._tip(
            marco,
            "Si el listado no trae AREA o TIPO, se usan estos valores.",
        ).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(16, 0))

    def _construir_rotaciones(self, padre):
        marco = self._card(
            padre,
            "Rotaciones a generar",
            "Usa Ctrl o Shift para seleccionar varias.",
            fila=0, columna=0, padx=(0, 7),
        )
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(0, weight=1)

        contenedor = tk.Frame(marco, bg=COLOR_TARJETA)
        contenedor.grid(row=0, column=0, sticky="nsew")
        contenedor.columnconfigure(0, weight=1)
        contenedor.rowconfigure(0, weight=1)

        self.lista_rot = tk.Listbox(
            contenedor, selectmode=tk.EXTENDED, height=12,
            exportselection=False, bg="#F8FAFC", fg=COLOR_TEXTO,
            selectbackground=COLOR_ACENTO_CLARO,
            selectforeground="#1E3A8A",
            relief="flat", borderwidth=0, highlightthickness=1,
            highlightbackground=COLOR_BORDE, highlightcolor=COLOR_ACENTO,
            font=(FUENTE, 10), activestyle="none",
        )
        self.lista_rot.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(
            contenedor, orient="vertical", command=self.lista_rot.yview,
            style="Vertical.TScrollbar",
        )
        barra.grid(row=0, column=1, sticky="ns")
        self.lista_rot.configure(yscrollcommand=barra.set)
        self.lista_rot.bind(
            "<<ListboxSelect>>", lambda _e: self._actualizar_conteo()
        )

        botones = tk.Frame(contenedor, bg=COLOR_TARJETA)
        botones.grid(row=0, column=2, sticky="ns", padx=(12, 0))
        ttk.Button(
            botones, text="Todas", style="Small.TButton", width=10,
            command=self._seleccionar_todas,
        ).grid(row=0, column=0, pady=(0, 6))
        ttk.Button(
            botones, text="Ninguna", style="Small.TButton", width=10,
            command=self._seleccionar_ninguna,
        ).grid(row=1, column=0)

        pie = tk.Frame(marco, bg=COLOR_TARJETA)
        pie.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.lbl_conteo = ttk.Label(
            pie, text="0 rotaciones seleccionadas", style="Count.TLabel"
        )
        self.lbl_conteo.pack(side="left")

    def _construir_opciones_outlook(self, padre):
        marco = self._card(
            padre,
            "Opciones de envio",
            "Alumnos en Para, tutor en CC. Sin envio automatico.",
            fila=0, columna=1, padx=(7, 0),
        )
        ttk.Label(
            marco, text="Modo de borrador", style="Card.TLabel"
        ).pack(anchor="w")
        self.combo_modo = ttk.Combobox(
            marco, textvariable=self.var_modo_borrador, state="readonly",
            values=[
                "Por grupo: presentacion + recordatorio",
                "Por grupo: solo presentacion",
                "Por grupo: solo recordatorio",
                "Por alumno: solo presentacion",
            ],
        )
        self.combo_modo.pack(fill="x", pady=(3, 8))
        ttk.Checkbutton(
            marco, text="Agregar al tutor en CC",
            variable=self.var_cc_tutor,
        ).pack(anchor="w", pady=2)
        ttk.Checkbutton(
            marco, text="Ocultar alumnos usando CCO",
            variable=self.var_cco,
        ).pack(anchor="w", pady=2)
        ttk.Checkbutton(
            marco, text="Mostrar el primer borrador al terminar",
            variable=self.var_mostrar,
        ).pack(anchor="w", pady=2)

        hora = tk.Frame(marco, bg=COLOR_TARJETA)
        hora.pack(fill="x", pady=(8, 4))
        ttk.Label(
            hora, text="Hora de ingreso", style="Card.TLabel"
        ).pack(side="left")
        ttk.Entry(
            hora, textvariable=self.var_hora_ingreso, width=8,
        ).pack(side="left", padx=(8, 0))

        ttk.Button(
            marco, text="Revisar grupos y tutores", style="Small.TButton",
            command=self._revisar_grupos,
        ).pack(anchor="w", pady=(10, 0))
        ttk.Label(
            marco,
            text="El correo del tutor se edita en el tercer Excel.",
            style="CardMuted.TLabel", wraplength=300, justify="left",
        ).pack(fill="x", pady=(8, 0))

    def _construir_progreso(self, padre):
        marco = self._card(
            padre,
            "Progreso y registro",
            "Sigue el avance y revisa el detalle.",
            fila=0, columna=0,
        )
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(2, weight=1)

        self.barra = ttk.Progressbar(
            marco, mode="determinate", maximum=1, value=0,
            style="Horizontal.TProgressbar",
        )
        self.barra.grid(row=0, column=0, sticky="ew")

        ttk.Label(
            marco, textvariable=self.var_estado, style="Card.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(10, 10))

        log_marco = tk.Frame(
            marco, bg="#0F172A", highlightbackground="#1E293B",
            highlightthickness=1,
        )
        log_marco.grid(row=2, column=0, sticky="nsew")
        log_marco.columnconfigure(0, weight=1)
        log_marco.rowconfigure(0, weight=1)

        self.registro = tk.Text(
            log_marco, height=14, state="disabled", wrap="word",
            bg="#0F172A", fg="#E2E8F0", insertbackground="#FFFFFF",
            relief="flat", borderwidth=0, highlightthickness=0,
            font=("Consolas", 9), padx=12, pady=10,
        )
        self.registro.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(
            log_marco, orient="vertical", command=self.registro.yview,
            style="Log.Vertical.TScrollbar",
        )
        barra.grid(row=0, column=1, sticky="ns")
        self.registro.configure(yscrollcommand=barra.set)

    def _construir_botones(self, padre):
        pie = tk.Frame(padre, bg=COLOR_FONDO)
        pie.grid(row=3, column=0, sticky="ew", padx=28, pady=(14, 18))
        pie.columnconfigure(0, weight=1)

        tk.Label(
            pie,
            text="Borradores sin envio automatico.",
            bg=COLOR_FONDO, fg=COLOR_TEXTO_SUAVE, font=(FUENTE, 9),
        ).grid(row=0, column=0, sticky="w")

        acciones = tk.Frame(pie, bg=COLOR_FONDO)
        acciones.grid(row=0, column=1, sticky="e")
        ttk.Button(
            acciones, text="Salir", style="Ghost.TButton",
            command=self._cerrar,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            acciones, text="Abrir carpeta", style="Ghost.TButton",
            command=self._abrir_carpeta,
        ).pack(side="left", padx=(0, 10))
        self.btn_outlook = ttk.Button(
            acciones, text="Borradores en Outlook",
            style="Secondary.TButton", command=self._crear_borradores,
        )
        self.btn_outlook.pack(side="left", padx=(0, 8))
        self.btn_planilla = ttk.Button(
            acciones, text="Planilla de correos",
            style="Secondary.TButton", command=self._generar_planilla,
        )
        self.btn_planilla.pack(side="left", padx=(0, 8))
        self.btn_generar = ttk.Button(
            acciones, text="Generar cartas", style="Accent.TButton",
            command=self._generar,
        )
        self.btn_generar.pack(side="left")

    # ------------------------------------------------------------------
    # SELECCION DE ARCHIVOS
    # ------------------------------------------------------------------
    def _elegir_excel(self, var_ruta, combo, preferencia, recargar):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
        )
        if not ruta:
            return
        var_ruta.set(ruta)
        try:
            hojas = g.listar_hojas(Path(ruta))
        except Exception as e:
            messagebox.showerror("No se pudo abrir el archivo", str(e))
            return

        combo["values"] = hojas
        elegida = hojas[0] if hojas else ""
        if preferencia:
            for hoja in hojas:
                if g.normalizar(preferencia) in g.normalizar(hoja):
                    elegida = hoja
                    break
        combo.set(elegida)

        if recargar:
            self._recargar_rotaciones()

    def _elegir_template(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar template de carta",
            filetypes=[("Documentos Word", "*.docx *.doc"), ("Todos", "*.*")],
        )
        if ruta:
            self.var_plantilla.set(ruta)

    def _elegir_salida(self):
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if ruta:
            self.var_salida.set(ruta)

    def _elegir_correo_template(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar plantilla de correo",
            filetypes=[("Documentos Word", "*.docx *.doc"), ("Todos", "*.*")],
        )
        if ruta:
            self.var_correo_template.set(ruta)

    def _elegir_correo_recordatorio(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar plantilla de recordatorio",
            filetypes=[("Documentos Word", "*.docx *.doc"), ("Todos", "*.*")],
        )
        if ruta:
            self.var_correo_recordatorio.set(ruta)

    def _elegir_comunes(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar archivos comunes para adjuntar"
        )
        if not rutas:
            return
        for ruta in rutas:
            if ruta not in self.rutas_comunes:
                self.rutas_comunes.append(ruta)
        self.var_comunes.set("; ".join(Path(r).name for r in self.rutas_comunes))

    def _quitar_comunes(self):
        self.rutas_comunes = []
        self.var_comunes.set("")

    def _revisar_grupos(self):
        lista = self.var_lista.get().strip()
        grilla_path = self.var_grilla.get().strip()
        if not lista or not Path(lista).exists():
            messagebox.showwarning(
                "Falta el listado", "Selecciona el listado de alumnos."
            )
            return
        if not grilla_path or not Path(grilla_path).exists():
            messagebox.showwarning(
                "Falta la planilla", "Selecciona la planilla de rotaciones."
            )
            return
        rotaciones = self._seleccionadas()
        if not rotaciones:
            messagebox.showwarning(
                "Sin rotaciones", "Selecciona al menos una rotacion."
            )
            return
        try:
            maestro = g.cargar_maestro(
                Path(lista), self.var_hoja_lista.get() or None
            )
            catalogo = self._catalogo_seleccionado()
        except Exception as e:
            messagebox.showerror("No se pudo preparar la revision", str(e))
            return
        fuente_actual = (grilla_path, self.var_hoja_grilla.get() or None)
        if self.grilla_fuente != fuente_actual or not self.grilla:
            self._recargar_rotaciones()
        if not self.grilla:
            messagebox.showerror(
                "Planilla invalida",
                "No se pudieron leer las rotaciones de la planilla seleccionada.",
            )
            return
        salida = self.var_salida.get().strip()
        base_salida = Path(salida) if salida else Path(lista).parent
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            _, grupos, _, _, _ = pc.construir_datos(
                self.grilla,
                maestro,
                rotaciones,
                base_salida,
                self.var_carrera.get().strip() or g.CARRERA,
                self.rutas_comunes,
                self.ultimos_generados,
                catalogo=catalogo,
                tipo=self.var_tipo.get().strip() or g.TIPO,
                area=self.var_area.get().strip() or g.AREA,
                hora_ingreso=self.var_hora_ingreso.get().strip() or "09:00",
            )
        except Exception as e:
            messagebox.showerror("No se pudieron revisar los grupos", str(e))
            return
        finally:
            self.root.config(cursor="")

        ventana = tk.Toplevel(self.root)
        ventana.title("Revision de grupos y tutores")
        ventana.geometry("1080x520")
        ventana.minsize(820, 360)
        ventana.configure(bg=COLOR_FONDO)

        ttk.Label(
            ventana,
            text=(
                "Alumnos en Para y tutor en CC. Corrige el correo del tutor "
                "en el tercer Excel y vuelve a revisar."
            ),
            style="Card.TLabel",
        ).pack(anchor="w", padx=16, pady=(14, 8))

        marco = tk.Frame(ventana, bg=COLOR_FONDO)
        marco.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(0, weight=1)

        columnas = (
            "Rotacion", "Campo", "Alumnos", "Tutor", "Correo_Tutor",
            "Cartas", "Revisar",
        )
        arbol = ttk.Treeview(
            marco, columns=columnas, show="headings", selectmode="browse"
        )
        anchos = {
            "Rotacion": 70, "Campo": 150, "Alumnos": 260, "Tutor": 150,
            "Correo_Tutor": 220, "Cartas": 60, "Revisar": 70,
        }
        for columna in columnas:
            arbol.heading(columna, text=columna)
            arbol.column(columna, width=anchos[columna], anchor="w")
        arbol.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(
            marco, orient="vertical", command=arbol.yview
        )
        barra.grid(row=0, column=1, sticky="ns")
        arbol.configure(yscrollcommand=barra.set)

        for grupo in grupos:
            arbol.insert("", "end", values=(
                grupo.get("Rotacion", ""),
                grupo.get("Campo_clinico", ""),
                grupo.get("Alumnos", ""),
                grupo.get("Tutor", ""),
                grupo.get("Correo_Tutor", ""),
                "SI" if grupo.get("Adjuntos_Cartas") else "NO",
                "SI" if grupo.get("Revisar_Tutor") else "NO",
            ))

        ttk.Button(
            ventana, text="Cerrar", style="Ghost.TButton",
            command=ventana.destroy,
        ).pack(anchor="e", padx=16, pady=(0, 14))

    def _cargar_hojas_iniciales(self):
        for var_ruta, combo, preferencia in (
            (self.var_lista, self.combo_hoja_lista, g.HOJA_MAESTRO),
            (self.var_grilla, self.combo_hoja_grilla, None),
            (self.var_catalogo, self.combo_hoja_catalogo, None),
        ):
            ruta = var_ruta.get().strip()
            if not ruta or not Path(ruta).exists():
                continue
            try:
                hojas = g.listar_hojas(Path(ruta))
            except Exception:
                continue
            combo["values"] = hojas
            elegida = hojas[0] if hojas else ""
            if preferencia:
                for hoja in hojas:
                    if g.normalizar(preferencia) in g.normalizar(hoja):
                        elegida = hoja
                        break
            combo.set(elegida)
        self._recargar_rotaciones()

    # ------------------------------------------------------------------
    # ROTACIONES
    # ------------------------------------------------------------------
    def _recargar_rotaciones(self):
        self.lista_rot.delete(0, tk.END)
        self.codigos_rot = []
        self.grilla = {}

        ruta = self.var_grilla.get().strip()
        if not ruta or not Path(ruta).exists():
            self.grilla_fuente = None
            self._actualizar_conteo()
            return

        hoja = self.var_hoja_grilla.get() or None
        try:
            self.grilla = g.parsear_grilla(Path(ruta), hoja)
            self.grilla_fuente = (ruta, hoja)
        except Exception as e:
            self._agregar_log(f"No se pudo leer la planilla de rotaciones: {e}")
            self.grilla_fuente = None
            self._actualizar_conteo()
            return

        for rot in g.ordenar_rotaciones(self.grilla):
            info = self.grilla[rot]
            total_alumnos = sum(len(v) for v in info["campos"].values())
            texto = f"{rot}    {info['fecha_str'] or 'sin fecha'}    ({total_alumnos} alumnos)"
            self.lista_rot.insert(tk.END, texto)
            self.codigos_rot.append(rot)

        self._seleccionar_todas()
        self._agregar_log(
            f"Planilla de rotaciones cargada: {len(self.codigos_rot)} rotaciones disponibles."
        )

    def _seleccionadas(self):
        return [self.codigos_rot[i] for i in self.lista_rot.curselection()]

    def _seleccionar_todas(self):
        if self.codigos_rot:
            self.lista_rot.selection_set(0, tk.END)
        self._actualizar_conteo()

    def _seleccionar_ninguna(self):
        self.lista_rot.selection_clear(0, tk.END)
        self._actualizar_conteo()

    def _actualizar_conteo(self):
        n = len(self._seleccionadas())
        self.lbl_conteo.config(text=f"{n} rotaciones seleccionadas")

    # ------------------------------------------------------------------
    # GENERACION
    # ------------------------------------------------------------------
    def _generar(self):
        if self.ejecutando:
            return

        lista = self.var_lista.get().strip()
        grilla = self.var_grilla.get().strip()
        template = self.var_plantilla.get().strip()
        salida = self.var_salida.get().strip()

        for nombre, ruta in (
            ("listado de alumnos", lista),
            ("planilla de rotaciones", grilla),
            ("template de carta", template),
        ):
            if not ruta or not Path(ruta).exists():
                messagebox.showwarning(
                    "Falta un archivo",
                    f"Selecciona el {nombre} antes de continuar.",
                )
                return

        if not salida:
            messagebox.showwarning("Falta la salida", "Selecciona una carpeta de salida.")
            return

        rotaciones = self._seleccionadas()
        if not rotaciones:
            messagebox.showwarning(
                "Sin rotaciones", "Selecciona al menos una rotacion."
            )
            return

        try:
            maestro = g.cargar_maestro(Path(lista), self.var_hoja_lista.get() or None)
        except Exception as e:
            messagebox.showerror("No se pudo leer el listado", str(e))
            return

        if not maestro:
            messagebox.showwarning(
                "Listado vacio",
                "No se encontraron alumnos en la hoja seleccionada.",
            )
            return

        fuente_actual = (grilla, self.var_hoja_grilla.get() or None)
        if self.grilla_fuente != fuente_actual or not self.grilla:
            self._recargar_rotaciones()
        if not self.grilla:
            messagebox.showerror(
                "Planilla invalida",
                "No se pudieron leer las rotaciones de la grilla seleccionada.",
            )
            return

        total, _ = g.planificar(self.grilla, maestro, rotaciones)
        if total == 0:
            messagebox.showwarning(
                "Sin cartas por generar",
                "No se encontraron alumnos del listado en las rotaciones seleccionadas.",
            )
            return

        if not messagebox.askyesno(
            "Confirmar generacion",
            f"Se generaran {total} cartas en:\n{salida}\n\nDeseas continuar?",
        ):
            return

        self._preparar_ejecucion(total)

        parametros = {
            "lista": lista,
            "hoja_lista": self.var_hoja_lista.get() or None,
            "grilla": self.grilla,
            "rotaciones": list(rotaciones),
            "template": template,
            "salida": salida,
            "carrera": self.var_carrera.get().strip() or g.CARRERA,
            "tipo": self.var_tipo.get().strip() or g.TIPO,
            "area": self.var_area.get().strip() or g.AREA,
            "pdf": self.var_pdf.get(),
        }

        hilo = threading.Thread(target=self._trabajar, args=(parametros,), daemon=True)
        hilo.start()

    def _generar_planilla(self):
        lista = self.var_lista.get().strip()
        grilla_path = self.var_grilla.get().strip()
        salida = self.var_salida.get().strip()

        if not lista or not Path(lista).exists():
            messagebox.showwarning(
                "Falta el listado", "Selecciona el listado de alumnos."
            )
            return
        if not grilla_path or not Path(grilla_path).exists():
            messagebox.showwarning(
                "Falta la planilla", "Selecciona la planilla de rotaciones."
            )
            return

        rotaciones = self._seleccionadas()
        if not rotaciones:
            messagebox.showwarning(
                "Sin rotaciones", "Selecciona al menos una rotacion."
            )
            return

        try:
            maestro = g.cargar_maestro(Path(lista), self.var_hoja_lista.get() or None)
        except Exception as e:
            messagebox.showerror("No se pudo leer el listado", str(e))
            return
        if not maestro:
            messagebox.showwarning(
                "Listado vacio",
                "No se encontraron alumnos en la hoja seleccionada.",
            )
            return

        fuente_actual = (grilla_path, self.var_hoja_grilla.get() or None)
        if self.grilla_fuente != fuente_actual or not self.grilla:
            self._recargar_rotaciones()
        if not self.grilla:
            messagebox.showerror(
                "Planilla invalida",
                "No se pudieron leer las rotaciones de la planilla seleccionada.",
            )
            return

        base_salida = Path(salida) if salida else Path(lista).parent
        inicial = base_salida if base_salida.exists() else Path(lista).parent
        destino = filedialog.asksaveasfilename(
            title="Guardar planilla de correos",
            initialdir=str(inicial),
            initialfile="Planilla_Correos.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
        )
        if not destino:
            return

        correo_template = self.var_correo_template.get().strip()
        if correo_template and not Path(correo_template).exists():
            messagebox.showwarning(
                "Plantilla no encontrada",
                "La plantilla de correo no existe. Se generara la planilla sin validarla.",
            )
            correo_template = ""
        correo_recordatorio = self.var_correo_recordatorio.get().strip()
        if correo_recordatorio and not Path(correo_recordatorio).exists():
            messagebox.showwarning(
                "Recordatorio no encontrado",
                "La plantilla de recordatorio no existe. Se usara una plantilla simple.",
            )
            correo_recordatorio = ""
        try:
            catalogo = self._catalogo_seleccionado()
        except Exception as e:
            messagebox.showerror("No se pudo leer el catalogo de tutores", str(e))
            return

        self.root.config(cursor="watch")
        self.root.update_idletasks()
        self._agregar_log("")
        self._agregar_log("Generando planilla de correos...")

        try:
            resultado = pc.exportar_planilla(
                self.grilla,
                maestro,
                rotaciones,
                base_salida,
                Path(destino),
                self.var_carrera.get().strip() or g.CARRERA,
                self.rutas_comunes,
                self.ultimos_generados,
                Path(correo_template) if correo_template else None,
                Path(correo_recordatorio) if correo_recordatorio else None,
                catalogo=catalogo,
                tipo=self.var_tipo.get().strip() or g.TIPO,
                area=self.var_area.get().strip() or g.AREA,
                hora_ingreso=self.var_hora_ingreso.get().strip() or "09:00",
            )
        except Exception as e:
            messagebox.showerror("No se pudo generar la planilla", str(e))
            self._agregar_log(f"ERROR: {e}")
            return
        finally:
            self.root.config(cursor="")

        self._agregar_log(f"Planilla generada: {resultado['archivo']}")
        self._agregar_log(
            f"Rotaciones: {resultado['rotaciones']} | "
            f"Grupos: {resultado['grupos']} | "
            f"Registros de alumnos: {resultado['alumnos']}"
        )
        if resultado["rotaciones_sin_adjuntos"]:
            self._agregar_log(
                f"Rotaciones sin cartas detectadas: "
                f"{resultado['rotaciones_sin_adjuntos']}"
            )
        if resultado["grupos_sin_tutor"]:
            self._agregar_log(
                f"Grupos con tutor por revisar: {resultado['grupos_sin_tutor']}"
            )
        if resultado["campos_faltantes_template"]:
            self._agregar_log(
                "Campos del template sin columna (revise los alias): "
                + ", ".join(resultado["campos_faltantes_template"])
            )

        mensaje = (
            f"Planilla generada con {resultado['rotaciones']} rotaciones, "
            f"{resultado['grupos']} grupos y "
            f"{resultado['alumnos']} registros de alumnos.\n\n"
            f"Archivo:\n{resultado['archivo']}"
        )
        if resultado["rotaciones_sin_adjuntos"]:
            mensaje += (
                f"\n\nAviso: {resultado['rotaciones_sin_adjuntos']} rotaciones "
                "no tienen cartas detectadas en la carpeta de salida."
            )
        if resultado["grupos_sin_tutor"]:
            mensaje += (
                f"\n\nAviso: {resultado['grupos_sin_tutor']} grupos no tienen "
                "correo de tutor confirmado. Revise la hoja "
                "Correspondencia_Grupos."
            )
        if resultado["campos_faltantes_template"]:
            mensaje += (
                "\n\nCampos del template sin columna: "
                + ", ".join(resultado["campos_faltantes_template"])
            )

        if messagebox.askyesno(
            "Planilla generada", mensaje + "\n\nDeseas abrir la planilla?"
        ):
            try:
                os.startfile(str(resultado["archivo"]))
            except Exception as e:
                messagebox.showinfo("No se pudo abrir", str(e))

    def _catalogo_seleccionado(self):
        ruta = self.var_catalogo.get().strip()
        if not ruta:
            return []
        if not Path(ruta).exists():
            raise FileNotFoundError(
                "El catalogo de tutores seleccionado no existe."
            )
        registros = ct.leer_catalogo(
            Path(ruta), self.var_hoja_catalogo.get() or None
        )
        if not registros:
            raise ValueError(
                "No se encontraron tutores en la hoja del catalogo."
            )
        return registros

    def _crear_borradores(self):
        if self.ejecutando:
            return

        lista = self.var_lista.get().strip()
        grilla_path = self.var_grilla.get().strip()

        if not lista or not Path(lista).exists():
            messagebox.showwarning(
                "Falta el listado", "Selecciona el listado de alumnos."
            )
            return
        if not grilla_path or not Path(grilla_path).exists():
            messagebox.showwarning(
                "Falta la planilla", "Selecciona la planilla de rotaciones."
            )
            return

        rotaciones = self._seleccionadas()
        if not rotaciones:
            messagebox.showwarning(
                "Sin rotaciones", "Selecciona al menos una rotacion."
            )
            return

        try:
            maestro = g.cargar_maestro(
                Path(lista), self.var_hoja_lista.get() or None
            )
        except Exception as e:
            messagebox.showerror("No se pudo leer el listado", str(e))
            return
        if not maestro:
            messagebox.showwarning(
                "Listado vacio",
                "No se encontraron alumnos en la hoja seleccionada.",
            )
            return

        fuente_actual = (grilla_path, self.var_hoja_grilla.get() or None)
        if self.grilla_fuente != fuente_actual or not self.grilla:
            self._recargar_rotaciones()
        if not self.grilla:
            messagebox.showerror(
                "Planilla invalida",
                "No se pudieron leer las rotaciones de la planilla seleccionada.",
            )
            return

        correo_template = self.var_correo_template.get().strip()
        if correo_template and not Path(correo_template).exists():
            messagebox.showwarning(
                "Plantilla no encontrada",
                "La plantilla de correo no existe. Se usara un cuerpo simple.",
            )
            correo_template = ""
        correo_recordatorio = self.var_correo_recordatorio.get().strip()
        if correo_recordatorio and not Path(correo_recordatorio).exists():
            messagebox.showwarning(
                "Recordatorio no encontrado",
                "La plantilla de recordatorio no existe. Se usara un cuerpo simple.",
            )
            correo_recordatorio = ""
        try:
            catalogo = self._catalogo_seleccionado()
        except Exception as e:
            messagebox.showerror("No se pudo leer el catalogo de tutores", str(e))
            return

        salida = self.var_salida.get().strip()
        base_salida = Path(salida) if salida else Path(lista).parent
        modos = {
            "Por grupo: presentacion + recordatorio":
                "grupo_presentacion_recordatorio",
            "Por grupo: solo presentacion": "grupo_presentacion",
            "Por grupo: solo recordatorio": "grupo_recordatorio",
            "Por alumno: solo presentacion": "alumno_presentacion",
        }
        modo = modos.get(
            self.var_modo_borrador.get(),
            "grupo_presentacion_recordatorio",
        )

        total_grupos, _ = g.planificar(self.grilla, maestro, rotaciones)
        if modo == "grupo_presentacion_recordatorio":
            detalle = (
                f"{total_grupos} grupos y {total_grupos * 2} borradores "
                "(presentacion + recordatorio)"
            )
        elif modo.startswith("grupo"):
            detalle = f"{total_grupos} grupos y {total_grupos} borradores"
        else:
            detalle = "un borrador por alumno"

        if not messagebox.askyesno(
            "Crear borradores en Outlook",
            f"Se crearan {detalle} para {len(rotaciones)} rotaciones.\n"
            "Los correos quedaran en Borradores y NO se enviaran.\n\n"
            "Deseas continuar?",
        ):
            return

        self.notebook.select(self.tab_proceso)
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        self._agregar_log("")
        self._agregar_log("Creando borradores en Outlook...")

        try:
            resultado = pc.crear_borradores_outlook(
                self.grilla,
                maestro,
                rotaciones,
                base_salida,
                self.var_carrera.get().strip() or g.CARRERA,
                self.rutas_comunes,
                self.ultimos_generados,
                Path(correo_template) if correo_template else None,
                Path(correo_recordatorio) if correo_recordatorio else None,
                catalogo=catalogo,
                modo=modo,
                usar_cco=self.var_cco.get(),
                cc_tutor=self.var_cc_tutor.get(),
                mostrar=self.var_mostrar.get(),
                tipo=self.var_tipo.get().strip() or g.TIPO,
                area=self.var_area.get().strip() or g.AREA,
                hora_ingreso=self.var_hora_ingreso.get().strip() or "09:00",
                log=self._agregar_log,
            )
        except Exception as e:
            messagebox.showerror(
                "No se pudieron crear los borradores", str(e)
            )
            self._agregar_log(f"ERROR: {e}")
            return
        finally:
            self.root.config(cursor="")

        self._agregar_log(
            f"Borradores creados: {resultado['creados']} "
            f"para {resultado['grupos']} grupos"
        )
        if resultado["sin_correo"]:
            self._agregar_log(
                "Sin correo: " + ", ".join(resultado["sin_correo"])
            )
        if resultado["sin_adjuntos"]:
            self._agregar_log(
                "Sin adjuntos: " + ", ".join(resultado["sin_adjuntos"])
            )
        if resultado["sin_tutor"]:
            self._agregar_log(
                "Grupos sin correo de tutor: "
                + ", ".join(resultado["sin_tutor"])
            )
        for etiqueta, error in resultado["errores"]:
            self._agregar_log(f"ERROR {etiqueta}: {error}")

        mensaje = (
            f"Se crearon {resultado['creados']} borradores en Outlook "
            f"para {resultado['grupos']} grupos.\n"
            "No se envio ningun correo."
        )
        if resultado["sin_correo"]:
            mensaje += "\n\nSin correo: " + ", ".join(resultado["sin_correo"])
        if resultado["sin_adjuntos"]:
            mensaje += (
                "\n\nSin adjuntos: " + ", ".join(resultado["sin_adjuntos"])
            )
        if resultado["sin_tutor"]:
            mensaje += (
                "\n\nSin correo de tutor: "
                + ", ".join(resultado["sin_tutor"])
            )
        if resultado["errores"]:
            mensaje += (
                "\n\nErrores: "
                + ", ".join(etiqueta for etiqueta, _ in resultado["errores"])
            )
        messagebox.showinfo("Borradores creados", mensaje)

    def _preparar_ejecucion(self, total):
        self.ejecutando = True
        self.notebook.select(self.tab_proceso)
        self.btn_generar.config(state="disabled")
        self.registro.config(state="normal")
        self.registro.delete("1.0", tk.END)
        self.registro.config(state="disabled")
        self.barra["maximum"] = max(total, 1)
        self.barra["value"] = 0
        self.var_estado.set(f"Generando 0 de {total} cartas...")
        self._agregar_log(f"Iniciando generacion de {total} cartas.")

    def _trabajar(self, parametros):
        try:
            import pythoncom
        except ImportError:
            pythoncom = None
        if pythoncom:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pythoncom = None

        try:
            self._log("Preparando template...")
            template_path = g.asegurar_template_docx(
                Path(parametros["template"]), log=self._log
            )

            self._log("Cargando listado de alumnos...")
            maestro = g.cargar_maestro(
                Path(parametros["lista"]), parametros["hoja_lista"]
            )
            self._log(f"{len(maestro)} alumnos cargados.")

            Path(parametros["salida"]).mkdir(parents=True, exist_ok=True)

            total_generadas = 0
            faltantes = []
            todas_generadas = []
            for rotacion in parametros["rotaciones"]:
                generadas, no_encontradas = g.procesar_rotacion(
                    rotacion,
                    parametros["grilla"],
                    maestro,
                    Path(parametros["salida"]),
                    not parametros["pdf"],
                    template_path,
                    parametros["carrera"],
                    parametros["tipo"],
                    parametros["area"],
                    log=self._log,
                    on_carta=self._una_carta,
                    silenciar_pdf=True,
                )
                total_generadas += len(generadas)
                faltantes.extend(no_encontradas)
                todas_generadas.extend(generadas)

            self.cola.put((
                "fin",
                {
                    "generadas": total_generadas,
                    "faltantes": faltantes,
                    "salida": parametros["salida"],
                    "generados": todas_generadas,
                },
            ))
        except Exception as e:
            self.cola.put(("error", str(e)))
        finally:
            if pythoncom:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # ------------------------------------------------------------------
    # COLA DE MENSAJES Y REGISTRO
    # ------------------------------------------------------------------
    def _log(self, mensaje):
        self.cola.put(("log", mensaje))

    def _una_carta(self):
        self.cola.put(("progreso", 1))

    def _procesar_cola(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self._agregar_log(dato)
                elif tipo == "progreso":
                    self.barra["value"] = self.barra["value"] + dato
                    self.var_estado.set(
                        f"Generando {int(self.barra['value'])} de "
                        f"{int(self.barra['maximum'])} cartas..."
                    )
                elif tipo == "fin":
                    self._terminar(dato)
                elif tipo == "error":
                    self._fallar(dato)
        except queue.Empty:
            pass
        self.root.after(100, self._procesar_cola)

    def _agregar_log(self, mensaje):
        self.registro.config(state="normal")
        self.registro.insert(tk.END, str(mensaje) + "\n")
        self.registro.see(tk.END)
        self.registro.config(state="disabled")

    def _terminar(self, resultado):
        self.ejecutando = False
        self.btn_generar.config(state="normal")
        self.barra["value"] = self.barra["maximum"]
        self.ultimos_generados = resultado.get("generados", [])

        generadas = resultado["generadas"]
        faltantes = resultado["faltantes"]
        self.var_estado.set(f"Proceso terminado. Cartas generadas: {generadas}")

        self._agregar_log("")
        self._agregar_log("=" * 50)
        self._agregar_log(f"Proceso terminado. Cartas generadas: {generadas}")
        self._agregar_log(
            "Puedes generar la planilla de correos con el boton "
            "'Generar planilla de correos'."
        )
        if faltantes:
            self._agregar_log(f"Alumnos no encontrados: {len(faltantes)}")
            for rot, campo, nombre in faltantes:
                self._agregar_log(f"  [{rot}] {campo}: {nombre}")

        mensaje = (
            f"Se generaron {generadas} cartas.\n\n"
            f"Carpeta de salida:\n{resultado['salida']}"
        )
        if faltantes:
            mensaje += f"\n\nAlumnos no encontrados: {len(faltantes)} (ver registro)."

        if messagebox.askyesno(
            "Proceso terminado", mensaje + "\n\nDeseas abrir la carpeta de salida?"
        ):
            self._abrir_carpeta()

    def _fallar(self, mensaje):
        self.ejecutando = False
        self.btn_generar.config(state="normal")
        self.var_estado.set("Ocurrio un error. Revisa el registro.")
        self._agregar_log("")
        self._agregar_log(f"ERROR: {mensaje}")
        messagebox.showerror("No se pudieron generar las cartas", mensaje)

    # ------------------------------------------------------------------
    # ACCIONES FINALES
    # ------------------------------------------------------------------
    def _abrir_carpeta(self):
        ruta = self.var_salida.get().strip()
        if ruta and Path(ruta).exists():
            os.startfile(ruta)
        else:
            messagebox.showinfo("Carpeta no disponible", "La carpeta de salida aun no existe.")

    def _cerrar(self):
        if self.ejecutando:
            if not messagebox.askyesno(
                "Proceso en curso",
                "Todavia se estan generando cartas. Deseas salir de todas formas?",
            ):
                return
        self.root.destroy()


def _activar_dpi():
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def main():
    _activar_dpi()
    root = tk.Tk()
    Aplicacion(root)
    root.mainloop()


if __name__ == "__main__":
    main()
