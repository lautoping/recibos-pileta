import os
import sys
import tempfile
import unittest
from datetime import date

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from utils import formatear_moneda, formatear_nombre_propio, parsear_fecha, sumar_un_mes, quitar_tildes
from alumnos_service import extraer_dias_clase
from excel_manager import crear_excel_precios_por_defecto, cargar_tabla_precios


class TestUtils(unittest.TestCase):
    def test_formatear_moneda(self):
        self.assertEqual(formatear_moneda(73000), "$73.000")
        self.assertEqual(formatear_moneda(1234.5), "$1.234,50")
        self.assertEqual(formatear_moneda("no es un número"), "$0")

    def test_formatear_nombre_propio(self):
        self.assertEqual(formatear_nombre_propio("juan   perez"), "Juan Perez")
        self.assertEqual(formatear_nombre_propio("gómez, maría"), "Gómez, María")

    def test_parsear_fecha(self):
        self.assertEqual(parsear_fecha("05/03/2026"), date(2026, 3, 5))
        self.assertEqual(parsear_fecha("05/03/26"), date(2026, 3, 5))
        self.assertIsNone(parsear_fecha("cualquier cosa"))

    def test_sumar_un_mes(self):
        self.assertEqual(sumar_un_mes(date(2026, 1, 31)), date(2026, 2, 28))
        self.assertEqual(sumar_un_mes(date(2026, 12, 15)), date(2027, 1, 15))

    def test_quitar_tildes(self):
        self.assertEqual(quitar_tildes("  Miércoles "), "miercoles")


class TestDiasClase(unittest.TestCase):
    def test_formatos_variados(self):
        self.assertEqual(extraer_dias_clase("Lu, Mi y Vi 15:30"), "Lunes, Miércoles, Viernes")
        self.assertEqual(extraer_dias_clase("martes y jueves 18hs"), "Martes, Jueves")
        self.assertEqual(extraer_dias_clase("sábados 10:00"), "Sábado")

    def test_vacio(self):
        self.assertEqual(extraer_dias_clase(""), "")
        self.assertEqual(extraer_dias_clase(None), "")


class TestPrecios(unittest.TestCase):
    def test_lee_la_tabla_por_defecto(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = os.path.join(carpeta, "precios.xlsx")
            crear_excel_precios_por_defecto(ruta)
            precios = cargar_tabla_precios(ruta)
        self.assertEqual(precios["inscripcion"], 15000)
        self.assertEqual(precios["carnet"], 1000)
        self.assertEqual(precios["cuotas"][2]["Ninguno"], 73000)
        self.assertEqual(precios["cuotas"][2]["Efectivo"], 66000)
        self.assertEqual(precios["cuotas"][3]["Grupo familiar"], 82000)


if __name__ == "__main__":
    unittest.main()
