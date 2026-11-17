import os
import sys
import tempfile
import unittest
import zipfile
from datetime import date

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from openpyxl import Workbook, load_workbook

from config import COLUMNAS_RECIBOS
from excel_manager import asegurar_excel_recibos, guardar_fila_recibo


class TestHojaSeleccionada(unittest.TestCase):
    """Al crear la hoja del mes nuevo, Excel no tiene que abrir las hojas agrupadas."""

    def test_solo_la_hoja_del_ultimo_mes_queda_seleccionada(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = os.path.join(tmp, "recibos.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "Octubre 2026"
            for i, col in enumerate(COLUMNAS_RECIBOS, start=1):
                ws.cell(row=1, column=i, value=col)
            ws.sheet_view.tabSelected = True
            wb.save(ruta)

            asegurar_excel_recibos(ruta, date(2026, 11, 2))
            fila = dict.fromkeys(COLUMNAS_RECIBOS, "")
            fila.update({"Nro Recibo": 1, "Fecha Emisión": "02/11/2026", "Nombre": "A"})
            guardar_fila_recibo(ruta, fila)

            wb = load_workbook(ruta)
            self.assertEqual(wb.active.title, "Noviembre 2026")
            seleccionadas = [w.title for w in wb.worksheets if w.sheet_view.tabSelected]
            self.assertEqual(seleccionadas, ["Noviembre 2026"])
            with zipfile.ZipFile(ruta) as z:
                xml = "".join(z.read(n).decode() for n in z.namelist() if n.startswith("xl/worksheets/sheet"))
            self.assertEqual(xml.count('tabSelected="1"'), 1)


if __name__ == "__main__":
    unittest.main()
