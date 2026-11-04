# Generador de Cartas

Aplicacion de escritorio para generar cartas de presentacion de practicas
clinicas, preparar la planilla de correos y crear borradores en Outlook.

La interfaz permite seleccionar el listado de alumnos, la planilla de
rotaciones, el catalogo de tutores y las plantillas Word. Tambien permite
revisar los grupos antes de generar, editar el diccionario de campos clinicos
y ajustar la configuracion sin modificar el codigo.

## Requisitos

- Windows 10 u 11.
- Python 3.11 o superior. El proyecto fue probado con Python 3.15.
- Tkinter incluido con Python en Windows.
- Microsoft Word instalado para convertir cartas a PDF.
- Microsoft Outlook instalado para crear los borradores.

## Instalacion

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Para desarrollo o empaquetado:

```bat
pip install -r requirements-dev.txt
```

## Ejecucion

Interfaz grafica:

```bat
Iniciar_interfaz.bat
```

o directamente:

```bat
python interfaz_cartas.py
```

Flujo recomendado en la interfaz:

1. Seleccionar el listado de alumnos.
2. Seleccionar la planilla de rotaciones.
3. Seleccionar el catalogo de tutores.
4. Elegir la plantilla de carta y las plantillas de correo.
5. Marcar las rotaciones y usar `Revisar grupos y tutores`.
6. Generar las cartas.
7. Generar la planilla de correos.
8. Crear los borradores de Outlook.

Los borradores se guardan en Outlook; la aplicacion no envia correos
automaticamente.

## Base de datos local

La aplicacion puede trabajar directamente con los Excel o con una base
SQLite local. La base se guarda por defecto en:

```text
data/cruce_correos.db
```

El boton `Base de datos` de la barra superior permite:

- Importar el listado de alumnos.
- Importar la planilla de rotaciones, indicando el periodo academico.
- Importar el catalogo de centros y tutores.
- Importar los tres archivos en una sola operacion.
- Activar la base como fuente de datos.
- Crear respaldos en `data/backups/`.
- Exportar las tablas principales a Excel.
- Revisar el historial de importaciones.

Cada rotacion guarda fecha de inicio, fecha de termino y el texto original.
La planilla de correos agrega las columnas `Fecha_Inicio` y `Fecha_Termino`,
manteniendo `fecha_rotacion` para compatibilidad.

La importacion es idempotente: volver a importar el mismo archivo actualiza
los registros existentes y no duplica alumnos, rotaciones ni asignaciones. Si
un alumno de la planilla no existe en el listado, se crea un registro
pendiente y se informa en las incidencias de la importacion.

### Flujo con base de datos

1. Abrir `Base de datos`.
2. Usar `Importar todo` con los archivos ya seleccionados, o importar cada
   archivo por separado.
3. Indicar el periodo de las rotaciones, por ejemplo `2026-S2`.
4. Marcar `Usar la base de datos como fuente de datos`.
5. Revisar y generar las cartas, la planilla y los borradores como siempre.

Tambien se puede volver al modo Excel desmarcando la misma opcion.

## Archivos de entrada

- Listado de alumnos: Excel con RUT, nombre y correo institucional.
- Planilla de rotaciones: Excel con columnas `R1`, `R2`, etc., campos
  clinicos y alumnos asignados.
- Catalogo de tutores: Excel con columnas como `Campo Clinico`, `Direccion`,
  `Contacto` y `Tutor(a)`.
- `Correo_Tutor` es opcional. Si no existe o esta vacia, el correo se busca
  dentro de `Tutor(a)` y `Contacto` usando una expresion regular.
- Plantillas Word: carta de presentacion, correo de presentacion y correo de
  recordatorio.

Los archivos `.xlsx` de entrada no se versionan porque son datos locales. Las
plantillas Word necesarias para ejecutar y empaquetar si se versionan.

## Salidas

- Cartas en `cartas_output/<rotacion>_<fecha>/`.
- Planilla Excel de correspondencia.
- Borradores de Outlook por grupo o por alumno, segun la configuracion.

## Configuracion

La configuracion se guarda en `config.json`:

- `data_source`: fuente activa, `excel` o `db`.
- `database`: motor y ruta de la base de datos.
- `campos_clinicos`: nombres completos y abreviaturas de campos clinicos.
- `alias_campos`: equivalencias entre nombres de campos.
- `outlook`: modo de borradores, uso de CCO y copia al tutor.
- `recordatorio`: plantilla y hora de ingreso.

El diccionario de campos clinicos tambien se puede editar desde el boton
`Diccionario` de la interfaz. La ruta de la base de datos se puede cambiar
manualmente en `database.path`; si es relativa, se resuelve desde la carpeta
de la aplicacion.

## Uso por linea de comandos

Generar todas las rotaciones:

```bat
python generar_cartas.py --all --lista ruta\lista.xlsx --grilla ruta\rotaciones.xlsx --template carta_template.docx
```

Generar rotaciones especificas y la planilla de correos:

```bat
python generar_cartas.py --rotacion R1 R2 --lista ruta\lista.xlsx --grilla ruta\rotaciones.xlsx --planilla
```

Opciones utiles:

- `--solo-docx`: genera Word sin convertir a PDF.
- `--hoja-lista` y `--hoja-grilla`: seleccionan una hoja del Excel.
- `--planilla-salida`: cambia la ruta del Excel de salida.
- `--adjuntos-comunes`: agrega archivos comunes a los correos.

## Pruebas

```bat
python -m unittest discover -s tests -v
```

Las pruebas cubren el parseo de rangos de fechas, la importacion a SQLite,
la idempotencia y la equivalencia entre los datos leidos desde Excel y desde
la base de datos.

## Empaquetado

```bat
pip install -r requirements-dev.txt
pyinstaller GeneradorCartas.spec
```

El ejecutable queda en `dist/GeneradorCartas/`. Las carpetas `build/`,
`dist/`, `cartas_output/`, `data/` y `__pycache__/` son generadas y no se
versionan. La base de datos se crea al iniciar la aplicacion; si se elimina
la carpeta `dist/`, conviene restaurar la base desde `data/backups/`.

## Estructura del proyecto

- `interfaz_cartas.py`: interfaz grafica Tkinter.
- `generar_cartas.py`: lectura de Excel, generacion de cartas y conversion.
- `planilla_correos.py`: planilla Excel y borradores de Outlook.
- `catalogo_tutores.py`: lectura y busqueda de tutores.
- `config.py`: carga y guardado de `config.json`.
- `datos/`: esquema, migraciones, importadores y repositorio SQLite.
- `data/`: base de datos local y respaldos.
- `tests/`: pruebas automatizadas.
- `GeneradorCartas.spec`: configuracion de PyInstaller.
- `requirements.txt`: dependencias de ejecucion.
- `requirements-dev.txt`: dependencias de desarrollo y empaquetado.

## Problemas comunes

- Si no se genera el PDF, verificar que Microsoft Word este instalado y que
  `docx2pdf` este disponible.
- Si no se crean los borradores, verificar que Outlook este instalado y que
  `pywin32` este correctamente instalado.
- Si un tutor aparece como `Sin coincidencia`, revisar que el campo clinico o
  el nombre del tutor coincidan con el catalogo.
- Si no aparece ningun correo de tutor, confirmar que el catalogo tenga las
  columnas `Tutor(a)` y `Contacto`, o agregar la columna opcional
  `Correo_Tutor`.
