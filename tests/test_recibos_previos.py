import os
import sys
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

import pandas as pd

from alumnos_service import extraer_dias_clase, extraer_datos_alumnos_dataframe
from excel_manager import buscar_cda_previa, calcular_vto_desde_recibo_anterior

RESUMEN = [
    {"nro": "297", "fecha": "01/10/2026", "nombre": "Jorge Cinquemani", "actividad": "Aquagym", "vto": "30/10/2026"},
    {"nro": "R297", "fecha": "03/10/2026", "nombre": "Jorge Cinquemani", "actividad": "Cobro de Saldo", "vto": "30/10/2026"},
    {"nro": "133", "fecha": "19/10/2026", "nombre": "Dante Cisterna", "actividad": "Clase Adaptación (CdA)", "vto": ""},
    {"nro": "150", "fecha": "20/09/2026", "nombre": "Ana Sola", "actividad": "Natación", "vto": "20/10/2026"},
    {"nro": "160", "fecha": "21/10/2026", "nombre": "Sin Vto", "actividad": "Natación", "vto": ""},
]


class TestRecibosPrevios(unittest.TestCase):
    def test_vto_anterior_mas_un_mes(self):
        vto, ant = calcular_vto_desde_recibo_anterior(RESUMEN, "Jorge  Cinquemani", "05/11/2026")
        self.assertEqual((vto, ant["nro"]), ("30/11/2026", "297"))

    def test_sin_recibo_en_el_mes_anterior(self):
        self.assertEqual(calcular_vto_desde_recibo_anterior(RESUMEN, "Ana Sola", "05/11/2026"), (None, None))
        vto, ant = calcular_vto_desde_recibo_anterior(RESUMEN, "Sin Vto", "05/11/2026")
        self.assertIsNone(vto)
        self.assertEqual(ant["nro"], "160")

    def test_cda_previa(self):
        self.assertEqual(buscar_cda_previa(RESUMEN, "Dante Cisterna", "21/10/2026")["nro"], "133")
        self.assertIsNone(buscar_cda_previa(RESUMEN, "Dante Cisterna", "18/10/2026"))


class TestDiasYNombres(unittest.TestCase):
    def test_variantes_de_dias(self):
        self.assertEqual(extraer_dias_clase("martesy jueves 8:00 hs"), "Martes, Jueves")
        self.assertEqual(extraer_dias_clase("martes y juves 18:30 hs"), "Martes, Jueves")
        self.assertEqual(extraer_dias_clase("martes 18:30 y sabados 15:30hs "), "Martes, Sábado")

    def test_nombre_con_espacios_dobles_tiene_detalles(self):
        df = pd.DataFrame([[None, "malena rincon  (asma)", None, None, None, "martes 18:30 y sabados 15:30hs "]])
        nombres, detalles = extraer_datos_alumnos_dataframe(df)
        self.assertEqual(nombres, ["Malena Rincon (Asma)"])
        self.assertEqual(detalles["malena rincon (asma)"]["dias_clase"], "Martes, Sábado")


if __name__ == "__main__":
    unittest.main()
