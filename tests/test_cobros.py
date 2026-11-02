import os
import sys
import tempfile
import unittest
from datetime import date

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from openpyxl import Workbook, load_workbook

from config import COLUMNAS_RECIBOS
from utils import limpiar_monto
from excel_manager import (
    guardar_fila_recibo,
    guardar_filas_recibo,
    obtener_recibos_con_deuda,
    registrar_pago_pendiente,
    eliminar_recibo,
    obtener_estado_recibos,
    asegurar_excel_recibos,
    recibos_con_mismo_numero_en_el_mes,
)


def fila(nro, fecha, nombre, total, pagado, estado, actividad="Natación"):
    d = dict.fromkeys(COLUMNAS_RECIBOS, "")
    d.update({
        "Nro Recibo": nro, "Fecha Emisión": fecha, "Nombre": nombre, "Actividad": actividad,
        "Total": total, "Monto Pagado": pagado, "Saldo Pendiente": total - pagado, "Estado Pago": estado,
    })
    return d


class TestCobros(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ruta = os.path.join(self.tmp.name, "recibos.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def _deudas(self):
        return [(d["nro_recibo"], d["nombre"], d["saldo_pendiente"]) for d in obtener_recibos_con_deuda(self.ruta)]

    def test_pago_parcial_no_duplica_la_deuda(self):
        guardar_fila_recibo(self.ruta, fila(12, "01/10/2026", "Ana", 73000, 0, "Debe"))
        d = obtener_recibos_con_deuda(self.ruta)[0]
        registrar_pago_pendiente(self.ruta, 12, "Ana", "05/10/2026", 30000, "Efectivo",
                                 fila_excel_especifica=d["fila_excel"], hoja=d["hoja"])
        self.assertEqual(self._deudas(), [(12, "Ana", 43000.0)])

    def test_saldar_no_toca_otro_talonario(self):
        guardar_fila_recibo(self.ruta, fila(12, "01/10/2026", "Ana", 73000, 0, "Debe"))
        guardar_fila_recibo(self.ruta, fila(12, "20/10/2026", "Beto", 41000, 0, "Debe"))
        ana = [d for d in obtener_recibos_con_deuda(self.ruta) if d["nombre"] == "Ana"][0]
        registrar_pago_pendiente(self.ruta, 12, "Ana", "21/10/2026", 73000, "Efectivo",
                                 fila_excel_especifica=ana["fila_excel"], hoja=ana["hoja"])
        self.assertEqual(self._deudas(), [(12, "Beto", 41000.0)])

    def test_eliminar_solo_el_recibo_de_esa_fecha(self):
        guardar_fila_recibo(self.ruta, fila(5, "01/10/2026", "Ana", 73000, 73000, "Pagado"))
        guardar_fila_recibo(self.ruta, fila(5, "25/10/2026", "Beto", 41000, 41000, "Pagado"))
        self.assertEqual(eliminar_recibo(self.ruta, 5, "01/10/2026"), 1)
        ultimos = obtener_estado_recibos(self.ruta)[3]
        self.assertEqual([u["Nombre"] for u in ultimos], ["Beto"])

    def test_cobro_sin_saldar_total_deja_el_resto(self):
        guardar_fila_recibo(self.ruta, fila(7, "01/10/2026", "Ana", 50000, 0, "Debe"))
        d = obtener_recibos_con_deuda(self.ruta)[0]
        registrar_pago_pendiente(self.ruta, 7, "Ana", "02/10/2026", 20000, "Efectivo",
                                 fila_excel_especifica=d["fila_excel"], hoja=d["hoja"], saldar_total=False)
        self.assertEqual(self._deudas(), [(7, "Ana", 30000.0)])

    def test_recibo_grupal_se_salda_completo(self):
        guardar_filas_recibo(self.ruta, [
            dict(fila(9, "03/10/2026", "Ana", 0, 0, "Debe"), **{"Total": "", "Monto Pagado": "", "Saldo Pendiente": ""}),
            fila(9, "03/10/2026", "Beto", 116000, 0, "Debe"),
        ])
        d = obtener_recibos_con_deuda(self.ruta)[0]
        registrar_pago_pendiente(self.ruta, 9, "Beto", "04/10/2026", 116000, "Efectivo",
                                 fila_excel_especifica=d["fila_excel"], hoja=d["hoja"])
        self.assertEqual(self._deudas(), [])
        ws = load_workbook(self.ruta)["Octubre 2026"]
        self.assertEqual([ws.cell(row=r, column=17).value for r in (2, 3)], ["Pagado", "Pagado"])


class TestHojasPorMes(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ruta = os.path.join(self.tmp.name, "recibos.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def test_deuda_del_mes_pasado_sigue_y_el_cobro_va_al_mes_nuevo(self):
        guardar_fila_recibo(self.ruta, fila(400, "30/10/2026", "Ana", 73000, 0, "Debe"))
        guardar_fila_recibo(self.ruta, fila(401, "02/11/2026", "Beto", 41000, 41000, "Pagado"))
        wb = load_workbook(self.ruta)
        self.assertIn("Octubre 2026", wb.sheetnames)
        self.assertIn("Noviembre 2026", wb.sheetnames)
        self.assertLess(wb.sheetnames.index("Octubre 2026"), wb.sheetnames.index("Noviembre 2026"))

        deudas = obtener_recibos_con_deuda(self.ruta)
        self.assertEqual([(d["nombre"], d["hoja"]) for d in deudas], [("Ana", "Octubre 2026")])

        registrar_pago_pendiente(self.ruta, 400, "Ana", "03/11/2026", 73000, "Efectivo",
                                 fila_excel_especifica=deudas[0]["fila_excel"], hoja=deudas[0]["hoja"])
        self.assertEqual(obtener_recibos_con_deuda(self.ruta), [])
        ws_nov = load_workbook(self.ruta)["Noviembre 2026"]
        self.assertEqual(ws_nov.cell(row=3, column=4).value, "Cobro de Saldo")
        self.assertEqual(obtener_estado_recibos(self.ruta)[0], 402)

    def test_hoja_vieja_recibos_se_pasa_a_hojas_mensuales(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Recibos"
        ws.append(COLUMNAS_RECIBOS)
        ws.append([fila(1, "15/10/2026", "Ana", 73000, 73000, "Pagado")[c] for c in COLUMNAS_RECIBOS])
        ws.append([fila(2, "", "Sin fecha", 73000, 73000, "Pagado")[c] for c in COLUMNAS_RECIBOS])
        ws.append([fila(3, "01/09/2026", "Carla", 41000, 0, "Debe")[c] for c in COLUMNAS_RECIBOS])
        wb.save(self.ruta)

        asegurar_excel_recibos(self.ruta, fecha_referencia=date(2026, 11, 2))
        wb = load_workbook(self.ruta)
        self.assertEqual(wb.sheetnames, ["Septiembre 2026", "Octubre 2026", "Noviembre 2026"])
        self.assertEqual(wb["Octubre 2026"].max_row, 3)  # encabezado + Ana + fila sin fecha
        self.assertEqual([d["nombre"] for d in obtener_recibos_con_deuda(self.ruta)], ["Carla"])
        copias = os.listdir(os.path.join(self.tmp.name, "copias_seguridad"))
        self.assertEqual(len(copias), 1)

    def test_aviso_de_numero_repetido_en_el_mes(self):
        guardar_fila_recibo(self.ruta, fila(242, "29/10/2026", "Ana", 66000, 66000, "Pagado"))
        self.assertEqual(len(recibos_con_mismo_numero_en_el_mes(self.ruta, 242, "30/10/2026")), 1)
        self.assertEqual(recibos_con_mismo_numero_en_el_mes(self.ruta, 242, "01/11/2026"), [])


class TestMontos(unittest.TestCase):
    def test_limpiar_monto(self):
        self.assertEqual(float(limpiar_monto("$73.000")), 73000)
        self.assertEqual(float(limpiar_monto("1.234,50")), 1234.5)
        self.assertEqual(float(limpiar_monto("1234.50")), 1234.5)
        self.assertEqual(float(limpiar_monto("1234,5")), 1234.5)
        self.assertEqual(float(limpiar_monto("73000")), 73000)
        self.assertEqual(float(limpiar_monto("1.234.567")), 1234567)


if __name__ == "__main__":
    unittest.main()
