import os
import sys
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from openpyxl import Workbook
from openpyxl.styles import Border, Side

from alumnos_service import extraer_grupos_familiares, buscar_alumno_en_lista, indice_grupos
from excel_manager import recibos_del_grupo_para_dto, texto_recibo_gf

LINEA = Side(style="thin")


def hoja_grupos():
    """Imita la 2da hoja: grupos encerrados con bordes, número a veces en la 2da fila."""
    wb = Workbook()
    wb.active.title = "2026"
    ws = wb.create_sheet("Grupo Fliar")
    filas = [
        ("MICA", None, None, None),
        (1, "Marcelo Molina", "T", "padre"),
        (None, "Alejandra marano", "", "madre"),
        (None, "thiago molina", "B", "hijo"),
        (2, "Carlos Lopez", "T", "matrimonio"),
        (None, "Silvia Ibañez", "B", ""),
        (None, None, None, None),
        ("ALE", None, None, None),
        (None, "nataly buzeta", "T", ""),       # el número quedó en la fila de abajo
        (26, "aymee rosales", "B", "madre e hija"),
        (14, "Valentin, Felipe y Misael Godoy", "TB", "hermanos"),
    ]
    for r, (a, b, borde, nota) in enumerate(filas, start=1):
        ws.cell(row=r, column=1, value=a)
        c = ws.cell(row=r, column=2, value=b)
        ws.cell(row=r, column=5, value=nota or None)
        if borde:
            c.border = Border(top=LINEA if "T" in borde else Side(), bottom=LINEA if "B" in borde else Side())
    return wb


ALUMNOS = [
    "Marcelo Molina", "Alejandra Marano", "Thiago Salvador Molina", "Carlos Lopez",
    "Nataly Buzeta", "Aymee Florencia Rosales", "Valentin Godoy", "Felipe Godoy", "Misael Godoy",
    "Osvaldo Quiroga",
]


class TestGruposFamiliares(unittest.TestCase):
    def test_arma_los_grupos_por_bordes(self):
        grupos = extraer_grupos_familiares(hoja_grupos(), ALUMNOS)
        self.assertEqual([g["miembros"] for g in grupos], [
            ["Marcelo Molina", "Alejandra Marano", "Thiago Salvador Molina"],
            ["Carlos Lopez", "Silvia Ibañez"],
            ["Nataly Buzeta", "Aymee Florencia Rosales"],
            ["Valentin Godoy", "Felipe Godoy", "Misael Godoy"],
        ])
        self.assertEqual([g["seccion"] for g in grupos], ["MICA", "MICA", "ALE", "ALE"])
        self.assertEqual(grupos[2]["numero"], "26")

    def test_busca_al_alumno_con_errores_de_tipeo(self):
        self.assertEqual(buscar_alumno_en_lista("Osvlado Quiroga", ALUMNOS), "Osvaldo Quiroga")
        self.assertEqual(buscar_alumno_en_lista("aymee rosales", ALUMNOS), "Aymee Florencia Rosales")
        self.assertIsNone(buscar_alumno_en_lista("Juan Perez", ALUMNOS))

    def test_recibos_del_grupo_para_elegir(self):
        grupos = extraer_grupos_familiares(hoja_grupos(), ALUMNOS)
        self.assertEqual(indice_grupos(grupos)["alejandra marano"], 0)
        resumen = [
            {"nro": "182", "fecha": "23/10/2026", "nombre": "Marcelo Molina", "actividad": "Natación"},
            {"nro": "R90", "fecha": "01/11/2026", "nombre": "Alejandra Marano", "actividad": "Cobro de Saldo"},
            {"nro": "50", "fecha": "01/09/2026", "nombre": "Alejandra Marano", "actividad": "Natación"},
            # Recibo de grupo donde Thiago también estaba: no sirve para justificar su dto.
            {"nro": "328", "fecha": "02/10/2026", "nombre": "Alejandra Marano", "actividad": "Natación"},
            {"nro": "328", "fecha": "02/10/2026", "nombre": "Thiago Salvador Molina", "actividad": "Natación"},
            {"nro": "10", "fecha": "05/06/2026", "nombre": "Marcelo Molina", "actividad": "Natación"},
            {"nro": "7", "fecha": "01/11/2026", "nombre": "Osvaldo Quiroga", "actividad": "Natación"},
        ]
        otros = ["Marcelo Molina", "Alejandra Marano"]
        res = recibos_del_grupo_para_dto(resumen, "Thiago Salvador Molina", otros, "02/11/2026")
        self.assertEqual([(r["nombre"], r["nro"]) for r in res], [("Marcelo Molina", "182"), ("Alejandra Marano", "50")])
        self.assertEqual(texto_recibo_gf(res[0]), "GF Marcelo Molina R182 del 23/10")
        # Para Marcelo sí vale el R328 (él no estaba en ese recibo)
        res_m = recibos_del_grupo_para_dto(resumen, "Marcelo Molina", ["Alejandra Marano", "Thiago Salvador Molina"], "02/11/2026")
        self.assertIn(("Alejandra Marano", "328"), [(r["nombre"], r["nro"]) for r in res_m])


if __name__ == "__main__":
    unittest.main()
