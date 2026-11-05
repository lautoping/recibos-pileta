import os
import sys
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

import pandas as pd

from alumnos_service import extraer_dias_clase, extraer_datos_alumnos_dataframe
from excel_manager import (
    buscar_cda_previa,
    buscar_ultimo_recibo_cuota,
    calcular_vto_desde_recibo_anterior,
    clases_de_recibo,
    metodo_de_recibo,
)

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

    def test_paga_unos_dias_tarde_sin_saltear_mes(self):
        # pagó el 31/10 (vto 30/11) y vuelve el 02/12: el vto sigue la cadena
        resumen = [{"nro": "400", "fecha": "31/10/2026", "nombre": "Leo Paz", "actividad": "Natación", "vto": "30/11/2026"}]
        vto, ant = calcular_vto_desde_recibo_anterior(resumen, "Leo Paz", "02/12/2026")
        self.assertEqual((vto, ant["nro"]), ("30/12/2026", "400"))

    def test_dos_recibos_en_el_mismo_mes_usa_el_ultimo(self):
        resumen = [
            {"nro": "401", "fecha": "01/10/2026", "nombre": "Leo Paz", "actividad": "Natación", "vto": "01/11/2026"},
            {"nro": "402", "fecha": "31/10/2026", "nombre": "Leo Paz", "actividad": "Natación", "vto": "01/12/2026"},
        ]
        vto, ant = calcular_vto_desde_recibo_anterior(resumen, "Leo Paz", "30/11/2026")
        self.assertEqual((vto, ant["nro"]), ("01/01/2027", "402"))

    def test_salteo_un_mes_usa_el_general(self):
        # R150: vto 20/10 -> el nuevo sería 20/11, que el 25/11 ya pasó
        vto, ant = calcular_vto_desde_recibo_anterior(RESUMEN, "Ana Sola", "25/11/2026")
        self.assertIsNone(vto)
        self.assertEqual(ant["nro"], "150")
        self.assertEqual(calcular_vto_desde_recibo_anterior(RESUMEN, "Nadie", "05/11/2026"), (None, None))
        vto, ant = calcular_vto_desde_recibo_anterior(RESUMEN, "Sin Vto", "05/11/2026")
        self.assertIsNone(vto)
        self.assertEqual(ant["nro"], "160")

    def test_ultimo_recibo_cuota_trae_metodo_y_clases(self):
        resumen = RESUMEN + [
            {"nro": "310", "fecha": "02/10/2026", "nombre": "Ana Sola", "actividad": "Natación", "vto": "",
             "clases_sem": 3, "metodo_pago": "Efectivo"},
            {"nro": "R310", "fecha": "04/10/2026", "nombre": "Ana Sola", "actividad": "Cobro de Saldo", "vto": "",
             "clases_sem": "", "metodo_pago": "Débito"},
        ]
        ult = buscar_ultimo_recibo_cuota(resumen, "ana  sola", "05/11/2026")
        self.assertEqual(ult["nro"], "310")
        self.assertEqual(clases_de_recibo(ult), 3)
        self.assertEqual(metodo_de_recibo(ult), ("Efectivo", None))
        self.assertIsNone(buscar_ultimo_recibo_cuota(resumen, "Ana Sola", "01/09/2026"))

    def test_metodo_y_clases_raros(self):
        self.assertEqual(metodo_de_recibo({"metodo_pago": "Mixto ($ 10.000 Efectivo / $ 5.000 Débito)"}),
                         ("Mixto", "Efectivo + Débito"))
        self.assertEqual(metodo_de_recibo({"metodo_pago": "Mixto ($ 1 Transferencia / $ 2 Débito)"}),
                         ("Mixto", "Transferencia + Débito"))
        self.assertEqual(metodo_de_recibo({"metodo_pago": "debito"}), ("Débito", None))
        self.assertEqual(metodo_de_recibo({"metodo_pago": "No figura"}), (None, None))
        self.assertEqual(clases_de_recibo({"clases_sem": "2.0"}), 2)
        self.assertIsNone(clases_de_recibo({"clases_sem": ""}))
        self.assertIsNone(clases_de_recibo({}))

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
