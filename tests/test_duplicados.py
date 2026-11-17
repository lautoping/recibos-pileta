import os
import sys
import tempfile
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from openpyxl import load_workbook

from config import COLUMNAS_RECIBOS
from excel_manager import (
    guardar_filas_recibo,
    buscar_recibos_duplicados,
    rehacer_recibo,
    obtener_estado_recibos,
    eliminar_recibo,
    resumen_recibos,
)


def fila(nro, fecha, nombre, total, obs="", carnet=""):
    d = dict.fromkeys(COLUMNAS_RECIBOS, "")
    d.update({
        "Nro Recibo": nro, "Fecha Emisión": fecha, "Nombre": nombre, "Actividad": "Natación",
        "Total": total, "Monto Pagado": total if total != "" else "", "Saldo Pendiente": 0 if total != "" else "",
        "Estado Pago": "Pagado", "Observaciones": obs, "Nro Carnet": carnet,
    })
    return d


class TestDuplicados(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ruta = os.path.join(self.tmp.name, "recibos.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def test_grupo_familiar_no_es_repetido(self):
        guardar_filas_recibo(self.ruta, [
            fila(10, "05/11/2026", "Ana", "", "Recibo grupal (2 pers)"),
            fila(10, "05/11/2026", "Beto", 100, "Recibo grupal (2 pers)"),
        ])
        # Formato viejo: total en cada fila, pero las dos dicen "Recibo grupal"
        guardar_filas_recibo(self.ruta, [
            fila(11, "05/11/2026", "Caro", 60, "Recibo grupal (2 pers)"),
            fila(11, "05/11/2026", "Dani", 50, "Recibo grupal (2 pers)"),
        ])
        self.assertEqual(buscar_recibos_duplicados(self.ruta), [])

    def test_mismo_numero_mismo_dia_y_otro_dia(self):
        guardar_filas_recibo(self.ruta, [fila(20, "05/11/2026", "Ana", 100)])
        guardar_filas_recibo(self.ruta, [fila(20, "05/11/2026", "Beto", 90)])  # mismo día, seguido
        guardar_filas_recibo(self.ruta, [fila(30, "06/11/2026", "Caro", 80)])
        guardar_filas_recibo(self.ruta, [fila(30, "09/11/2026", "Dani", 70)])  # 3 días después
        dups = buscar_recibos_duplicados(self.ruta)
        self.assertEqual(sorted(d["nro"] for d in dups), ["20", "30"])
        d20 = next(d for d in dups if d["nro"] == "20")
        self.assertEqual([b["nombres"] for b in d20["recibos"]], [["Ana"], ["Beto"]])

    def test_otra_vuelta_del_talonario_no_es_repetido(self):
        guardar_filas_recibo(self.ruta, [fila(5, "01/10/2026", "Ana", 100)])
        guardar_filas_recibo(self.ruta, [fila(5, "05/11/2026", "Beto", 90)])
        self.assertEqual(buscar_recibos_duplicados(self.ruta), [])

    def test_anulado_no_es_repetido(self):
        guardar_filas_recibo(self.ruta, [fila(7, "05/11/2026", "Ana", 100, "Anulado")])
        guardar_filas_recibo(self.ruta, [fila(7, "05/11/2026", "Beto", 90)])
        self.assertEqual(buscar_recibos_duplicados(self.ruta), [])

    def test_rehacer_pasa_al_numero_nuevo_y_al_final(self):
        guardar_filas_recibo(self.ruta, [fila(20, "05/11/2026", "Ana", 100)])
        guardar_filas_recibo(self.ruta, [fila(20, "05/11/2026", "Beto", 90, "obs")])
        guardar_filas_recibo(self.ruta, [fila(21, "06/11/2026", "Caro", 80)])
        b = buscar_recibos_duplicados(self.ruta)[0]["recibos"][1]
        n = rehacer_recibo(self.ruta, b["hoja"], b["fila"], b["nro"], b["fecha"], 22, "07/11/2026")
        self.assertEqual(n, 1)
        ws = load_workbook(self.ruta).worksheets[-1]
        filas = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]
        self.assertEqual([(r[0], r[1], r[2]) for r in filas], [
            (20, "05/11/2026", "Ana"), (21, "06/11/2026", "Caro"), (22, "07/11/2026", "Beto"),
        ])
        self.assertIn("Rehecho: antes R20 del 05/11", filas[-1][17])
        self.assertTrue(filas[-1][17].startswith("obs"))
        self.assertEqual(buscar_recibos_duplicados(self.ruta), [])
        self.assertEqual(obtener_estado_recibos(self.ruta)[0], 23)

    def test_eliminar_no_se_lleva_el_repetido_de_al_lado(self):
        guardar_filas_recibo(self.ruta, [fila(20, "05/11/2026", "Ana", 100)])
        guardar_filas_recibo(self.ruta, [fila(20, "05/11/2026", "Beto", 90)])
        self.assertEqual(eliminar_recibo(self.ruta, 20, "05/11/2026", fila_ref=3), 1)
        ws = load_workbook(self.ruta).worksheets[-1]
        self.assertEqual(ws.cell(row=2, column=3).value, "Ana")

    def test_proximo_carnet_con_jubilado_primero(self):
        guardar_filas_recibo(self.ruta, [
            fila(10, "05/11/2026", "Jubilada", "", "Recibo grupal (2 pers)", carnet=173),
            fila(10, "05/11/2026", "Familiar", 124, "Recibo grupal (2 pers)", carnet=172),
        ])
        self.assertEqual(obtener_estado_recibos(self.ruta)[1], 174)

    def test_resumen_sin_poder_leer_devuelve_none(self):
        with open(self.ruta, "w") as f:
            f.write("no es un xlsx")
        self.assertIsNone(resumen_recibos(self.ruta, reintentos=2, espera=0))


if __name__ == "__main__":
    unittest.main()
