import os
import sys
import unittest
from datetime import date

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from utils import revisar_atraso, dias_semana_de_texto


class TestAtrasos(unittest.TestCase):
    def test_dias_de_texto(self):
        self.assertEqual(dias_semana_de_texto("Lunes, Miércoles"), {0, 2})
        self.assertEqual(dias_semana_de_texto("sábado"), {5})
        self.assertIsNone(dias_semana_de_texto(""))
        self.assertIsNone(dias_semana_de_texto("de 15 a 16"))

    def test_viene_despues_del_pago_en_termino(self):
        # Vence lunes 2/11/2026, viene los miércoles y paga el miércoles 4: en término
        self.assertEqual(date(2026, 11, 2).weekday(), 0)
        self.assertEqual(revisar_atraso("02/11/2026", "04/11/2026", "Miércoles"), ("en_termino", None))

    def test_clase_entre_vto_y_pago_atrasado(self):
        # Vence lunes 2, viene martes 3 y paga jueves 5: atrasado por la clase del 3
        self.assertEqual(
            revisar_atraso("02/11/2026", "05/11/2026", "Martes, Jueves"), ("atrasado", date(2026, 11, 3))
        )

    def test_dia_del_vto_y_dia_del_pago_no_cuentan(self):
        # Clase el lunes 2 (día del vto) y paga el lunes 9: no hubo otra clase en el medio
        self.assertEqual(revisar_atraso("02/11/2026", "09/11/2026", "Lunes"), ("en_termino", None))
        self.assertEqual(revisar_atraso("02/11/2026", "10/11/2026", "Lunes")[0], "atrasado")

    def test_paga_antes_del_vto(self):
        self.assertEqual(revisar_atraso("02/12/2026", "28/11/2026", ""), ("en_termino", None))

    def test_sin_datos(self):
        self.assertEqual(revisar_atraso("02/11/2026", "20/11/2026", ""), ("sin_dias", None))
        self.assertEqual(revisar_atraso("", "20/11/2026", "Lunes"), ("sin_fechas", None))


if __name__ == "__main__":
    unittest.main()
