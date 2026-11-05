"""Pruebas basicas de la capa de datos."""

import tempfile
import unittest
from pathlib import Path

import catalogo_tutores as ct
import generar_cartas as g
import planilla_correos as pc
from datos.fechas import parsear_rango_fechas, sugerir_periodo
from datos.repositorio import Repositorio

BASE = Path(__file__).resolve().parent.parent
LISTA = BASE / "Lista_Preclinicos_II_Kinesiologia.xlsx"
GRILLA = BASE / "Lista_Preclinicos_II_Kinesiologia_para_compartir_2.xlsx"
CATALOGO = BASE / "Test_Tutores.xlsx"


class FechasTest(unittest.TestCase):
    def test_rango_con_al(self):
        datos = parsear_rango_fechas("31-08-2026 al 04-09-2026")
        self.assertEqual(datos["fecha_inicio"], "2026-08-31")
        self.assertEqual(datos["fecha_fin"], "2026-09-04")
        self.assertEqual(
            datos["fecha_texto_normalizado"], "31-08-2026 al 04-09-2026"
        )

    def test_rango_sin_al(self):
        datos = parsear_rango_fechas("25-09-2026 28-10-2026")
        self.assertEqual(datos["fecha_inicio"], "2026-09-25")
        self.assertEqual(datos["fecha_fin"], "2026-10-28")

    def test_una_sola_fecha(self):
        datos = parsear_rango_fechas("02-11-2026")
        self.assertEqual(datos["fecha_inicio"], "2026-11-02")
        self.assertEqual(datos["fecha_fin"], "")

    def test_periodo(self):
        self.assertEqual(sugerir_periodo("2026-08-31"), "2026-S2")
        self.assertEqual(sugerir_periodo("2026-03-01"), "2026-S1")


class RepositorioTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.config = {
            "database": {
                "engine": "sqlite",
                "path": str(Path(self.temp.name) / "test.db"),
            }
        }
        self.repo = Repositorio(self.config)

    def tearDown(self):
        self.repo.cerrar()
        self.temp.cleanup()

    def test_importacion_completa(self):
        alumnos = self.repo.importar_alumnos(LISTA, None, "KINESIOLOGIA")
        rotaciones = self.repo.importar_rotaciones(
            GRILLA, None, "KINESIOLOGIA", "2026-S2"
        )
        catalogo = self.repo.importar_catalogo(CATALOGO, None, "KINESIOLOGIA")

        self.assertGreater(alumnos["insertados"], 0)
        self.assertEqual(rotaciones["insertados"], 15)
        self.assertGreater(catalogo["insertados"], 0)

        grilla = self.repo.obtener_grilla("KINESIOLOGIA", "2026-S2")
        self.assertIn("R1", grilla)
        self.assertEqual(grilla["R1"]["fecha_inicio"], "2026-08-31")
        self.assertEqual(grilla["R1"]["fecha_fin"], "2026-09-04")
        self.assertTrue(grilla["R1"]["campos"])

        maestro = self.repo.obtener_maestro("KINESIOLOGIA")
        self.assertGreater(len(maestro), 0)
        self.assertTrue(any(v.get("rut_fmt") for v in maestro.values()))

        catalogo_db = self.repo.obtener_catalogo("KINESIOLOGIA")
        self.assertTrue(any(r.get("correo_tutor") for r in catalogo_db))

    def test_importacion_idempotente(self):
        self.repo.importar_alumnos(LISTA, None, "KINESIOLOGIA")
        self.repo.importar_rotaciones(GRILLA, None, "KINESIOLOGIA", "2026-S2")
        self.repo.importar_catalogo(CATALOGO, None, "KINESIOLOGIA")
        primero = self.repo.resumen()

        self.repo.importar_alumnos(LISTA, None, "KINESIOLOGIA")
        self.repo.importar_rotaciones(GRILLA, None, "KINESIOLOGIA", "2026-S2")
        self.repo.importar_catalogo(CATALOGO, None, "KINESIOLOGIA")
        segundo = self.repo.resumen()

        for tabla in (
            "alumnos", "rotaciones", "asignaciones", "campos_clinicos",
            "centros_clinicos", "tutores", "catalogo_registros",
        ):
            self.assertEqual(primero[tabla], segundo[tabla], tabla)

    def test_equivalencia_excel_y_bd(self):
        self.repo.importar_alumnos(LISTA, None, "KINESIOLOGIA")
        self.repo.importar_rotaciones(GRILLA, None, "KINESIOLOGIA", "2026-S2")
        self.repo.importar_catalogo(CATALOGO, None, "KINESIOLOGIA")

        def grupos(maestro, grilla, catalogo):
            return pc.construir_datos(
                grilla, maestro, list(grilla), BASE / "cartas_output",
                "KINESIOLOGIA", catalogo=catalogo, tipo=g.TIPO, area=g.AREA,
            )[1]

        excel = grupos(
            g.cargar_maestro(LISTA),
            g.parsear_grilla(GRILLA),
            ct.leer_catalogo(CATALOGO),
        )
        bd = grupos(
            self.repo.obtener_maestro("KINESIOLOGIA"),
            self.repo.obtener_grilla("KINESIOLOGIA", "2026-S2"),
            self.repo.obtener_catalogo("KINESIOLOGIA"),
        )
        claves = (
            "Rotacion", "Campo_clinico", "Alumnos", "Tutor", "Correo_Tutor",
            "Correo_Tutor_Origen", "Direccion", "Contacto",
            "Nombre_Residencia", "Hora_Ingreso", "Revisar_Tutor",
            "Fecha_Inicio", "Fecha_Termino",
        )
        self.assertEqual(len(excel), len(bd))
        for esperado, obtenido in zip(excel, bd):
            for clave in claves:
                self.assertEqual(esperado.get(clave), obtenido.get(clave), clave)

    def test_exportar_excel(self):
        self.repo.importar_alumnos(LISTA, None, "KINESIOLOGIA")
        self.repo.importar_rotaciones(GRILLA, None, "KINESIOLOGIA", "2026-S2")
        self.repo.importar_catalogo(CATALOGO, None, "KINESIOLOGIA")
        destino = Path(self.temp.name) / "respaldo.xlsx"
        self.repo.exportar_excel(destino)
        self.assertTrue(destino.exists())

    def test_carreras_detectadas(self):
        catalogo = self.repo.importar_catalogo(CATALOGO, None, "KINESIOLOGIA")
        self.assertEqual(len(catalogo["carreras"]), 4)
        carreras = self.repo.listar_carreras()
        nombres = {g.normalizar(fila["nombre"]) for fila in carreras}
        self.assertEqual(
            nombres,
            {
                g.normalizar("Kinesiología"),
                g.normalizar("Enfermería"),
                g.normalizar("Fonoaudiología"),
                g.normalizar("Terapia Ocupacional"),
            },
        )
        total_catalogo = sum(fila["catalogo"] for fila in carreras)
        self.assertEqual(total_catalogo, 7)

        analisis = self.repo.analizar_alumnos(LISTA, None, "KINESIOLOGIA")
        self.assertEqual(len(analisis["carreras"]), 1)
        self.assertEqual(
            g.normalizar(analisis["carreras"][0]["nombre"]),
            g.normalizar("Kinesiología"),
        )

        self.repo.importar_alumnos(LISTA, None, "KINESIOLOGIA")
        alumnos = self.repo.listar_alumnos()
        self.assertTrue(alumnos)
        self.assertEqual(
            {g.normalizar(fila["carrera"]) for fila in alumnos},
            {g.normalizar("Kinesiología")},
        )

    def test_explorador_filtros(self):
        self.repo.importar_catalogo(CATALOGO, None, "KINESIOLOGIA")
        self.repo.importar_alumnos(LISTA, None, "KINESIOLOGIA")
        self.repo.importar_rotaciones(GRILLA, None, "KINESIOLOGIA", "2026-S2")

        centros = self.repo.listar_centros()
        self.assertTrue(centros)
        centro = centros[0]
        alumnos_centro = self.repo.listar_alumnos("", centro["id"])
        self.assertIsInstance(alumnos_centro, list)

        alumnos_sin_centro = self.repo.listar_alumnos("", -1)
        self.assertGreater(len(alumnos_sin_centro), 0)
        self.assertLessEqual(len(alumnos_sin_centro), len(self.repo.listar_alumnos()))
        self.assertEqual(self.repo.listar_centros("", -1), [])

        tutores = self.repo.listar_tutores("Kinesiología", centro["id"])
        self.assertIsInstance(tutores, list)

        catalogo = self.repo.listar_catalogo("Kinesiología")
        self.assertTrue(catalogo)
        self.assertTrue(
            all(
                g.normalizar(fila["carrera"]) == g.normalizar("Kinesiología")
                for fila in catalogo
            )
        )

        columnas = (
            ("carrera", "Carrera", 150),
            ("campo", "Campo clinico", 200),
            ("centro", "Centro clinico", 200),
            ("tutor", "Tutor", 180),
        )
        destino = Path(self.temp.name) / "vista.xlsx"
        self.repo.exportar_vista(destino, columnas, catalogo)
        self.assertTrue(destino.exists())


if __name__ == "__main__":
    unittest.main()
