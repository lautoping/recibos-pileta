"""
Genera datos de ejemplo (todos inventados) para probar la app sin datos reales.

Crea, en la carpeta de la app:
- precios.xlsx           tabla de precios con la estructura estándar
- alumnos_ejemplo.xlsx   planilla de alumnos con el mismo formato que la planilla real
- cache_alumnos.json     la lista de alumnos ya cargada (como si se hubiera sincronizado)
- recibos.xlsx           algunos recibos de ejemplo, con pagos completos, parciales y deudas
- config_recibos.json    configuración apuntando a esos archivos

No pisa archivos que ya existan, para no borrar datos reales por error.
Para regenerarlos igual: python generar_datos_ejemplo.py --forzar
"""

import os
import sys
import json
import random
import argparse
from datetime import date, timedelta

import pandas as pd

DIR_ACTUAL = os.path.dirname(os.path.abspath(__file__))
if DIR_ACTUAL not in sys.path:
    sys.path.insert(0, DIR_ACTUAL)

from config import (
    ARCHIVO_CONFIG,
    ARCHIVO_CACHE_ALUMNOS,
    ARCHIVO_PRECIOS_DEFAULT,
    ARCHIVO_RECIBOS_DEFAULT,
)
from utils import sumar_un_mes, quitar_tildes
from excel_manager import crear_excel_precios_por_defecto, guardar_fila_recibo, cargar_tabla_precios
from alumnos_service import extraer_datos_alumnos_dataframe, guardar_cache_alumnos

ARCHIVO_ALUMNOS_EJEMPLO = os.path.join(DIR_ACTUAL, "alumnos_ejemplo.xlsx")

NOMBRES = [
    "Martina", "Sofía", "Valentina", "Camila", "Lucía", "Julieta", "Florencia", "Agustina",
    "Carolina", "Paula", "Mariana", "Gabriela", "Silvia", "Norma", "Graciela", "Susana",
    "Tomás", "Joaquín", "Santiago", "Mateo", "Facundo", "Nicolás", "Federico", "Lucas",
    "Martín", "Diego", "Pablo", "Ricardo", "Jorge", "Osvaldo", "Héctor", "Raúl",
]
APELLIDOS = [
    "Ejemplar", "Prueba", "Muestra", "Ficticio", "Inventado", "Demo", "Modelo", "Ensayo",
    "Testeo", "Simulado", "Ilustrativo", "Provisorio",
]
HORARIOS = [
    "lunes y miércoles 9:00 a 10:00",
    "Lu, Mi y Vi 15:30",
    "martes y jueves 18hs",
    "Ma y Ju 19:00 a 20:00",
    "sábados 10:00",
    "lunes, miércoles y viernes 8:00",
    "miercoles 17:00",
    "Vie 20hs",
]


def generar_alumnos(rnd, cantidad=48):
    """Planilla con el formato de la real: nombre en la col. 1, jubilado en la 2, horario en la 5."""
    filas = [["N°", "Apellido y nombre", "Condición", "Teléfono", "Fecha alta", "Días y horario"]]
    usados = set()
    while len(filas) <= cantidad:
        nombre = f"{rnd.choice(NOMBRES)} {rnd.choice(APELLIDOS)}"
        if nombre in usados:
            continue
        usados.add(nombre)
        jubilado = "Jubilado/a" if rnd.random() < 0.18 else ""
        filas.append([
            len(filas),
            nombre,
            jubilado,
            "",
            "",
            rnd.choice(HORARIOS),
        ])
    return pd.DataFrame(filas)


def generar_recibos(rnd, nombres, detalles, precios, desde, hasta, cantidad=40):
    """Recibos variados: efectivo, transferencia, mixto, parciales, deudas, inscripciones."""
    inscripcion = int(precios["inscripcion"])

    def precio(clases_semana, descuento):
        return int(precios["cuotas"][clases_semana][descuento])

    recibos = []
    fecha = desde
    nro_recibo = 1
    nro_carnet = 1
    for _ in range(cantidad):
        fecha = min(fecha + timedelta(days=rnd.choice([0, 0, 1, 1, 2])), hasta)
        nombre = rnd.choice(nombres)
        det = detalles.get(quitar_tildes(nombre), {})
        actividad = rnd.choice(["Natación"] * 5 + ["Aquagym"] * 2)
        clases = rnd.choice([1, 2, 2, 2, 3])
        metodo = rnd.choice(["Efectivo"] * 5 + ["Transferencia"] * 3 + ["Débito"])

        if det.get("es_jubilado"):
            descuento = "Jubilado"
        elif rnd.random() < 0.15:
            descuento = "Grupo familiar"
        elif metodo == "Efectivo":
            descuento = "Efectivo"
        else:
            descuento = "Ninguno"

        cuota = precio(clases, descuento)
        es_nuevo = rnd.random() < 0.12
        monto_insc = inscripcion if es_nuevo else 0
        total = cuota + monto_insc

        estado = "Pagado"
        pagado = total
        if rnd.random() < 0.10:
            estado, pagado = "Parcial", int(round(total * 0.5, -3))
        elif rnd.random() < 0.05:
            estado, pagado = "Debe", 0

        if metodo == "Transferencia" and rnd.random() < 0.2 and pagado:
            efectivo = int(round(pagado * 0.4, -3))
            metodo = f"Mixto (${efectivo:,} Efectivo / ${pagado - efectivo:,} Transferencia)".replace(",", ".")

        carnet = ""
        if es_nuevo:
            carnet = nro_carnet
            nro_carnet += 1

        recibos.append({
            "Nro Recibo": nro_recibo,
            "Fecha Emisión": fecha.strftime("%d/%m/%Y"),
            "Nombre": nombre,
            "Actividad": actividad,
            "Detalle Actividad": "",
            "Clases por Semana": clases,
            "Fecha Vencimiento": sumar_un_mes(fecha).strftime("%d/%m/%Y"),
            "Inscripción": "Sí" if es_nuevo else "No",
            "Descuento": descuento,
            "Método de Pago": metodo if pagado else "",
            "Nro Carnet": carnet,
            "Monto Cuota": cuota,
            "Monto Inscripción": monto_insc,
            "Total": total,
            "Monto Pagado": pagado,
            "Saldo Pendiente": total - pagado,
            "Estado Pago": estado,
            "Observaciones": "Datos de ejemplo",
        })
        nro_recibo += 1
    return recibos


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--forzar", action="store_true", help="regenerar aunque los archivos ya existan")
    parser.add_argument("--semilla", type=int, default=7, help="semilla para que los datos salgan siempre iguales")
    args = parser.parse_args()

    archivos = [ARCHIVO_PRECIOS_DEFAULT, ARCHIVO_ALUMNOS_EJEMPLO, ARCHIVO_CACHE_ALUMNOS,
                ARCHIVO_RECIBOS_DEFAULT, ARCHIVO_CONFIG]
    existentes = [a for a in archivos if os.path.exists(a)]
    if existentes and not args.forzar:
        print("Ya existen estos archivos y no los voy a pisar (pueden ser datos reales):")
        for a in existentes:
            print(f"  - {os.path.basename(a)}")
        print("Si querés regenerarlos igual: python generar_datos_ejemplo.py --forzar")
        sys.exit(1)

    if args.forzar:
        for a in existentes:
            os.replace(a, a + ".bak")

    rnd = random.Random(args.semilla)

    crear_excel_precios_por_defecto(ARCHIVO_PRECIOS_DEFAULT)
    precios = cargar_tabla_precios(ARCHIVO_PRECIOS_DEFAULT)

    df_alumnos = generar_alumnos(rnd)
    df_alumnos.to_excel(ARCHIVO_ALUMNOS_EJEMPLO, header=False, index=False)
    nombres, detalles = extraer_datos_alumnos_dataframe(df_alumnos)
    guardar_cache_alumnos(nombres, detalles)

    hoy = date.today()
    desde = hoy - timedelta(days=45)
    for fila in generar_recibos(rnd, nombres, detalles, precios, desde, hasta=hoy):
        guardar_fila_recibo(ARCHIVO_RECIBOS_DEFAULT, fila)

    with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as f:
        json.dump({
            "archivo_recibos": ARCHIVO_RECIBOS_DEFAULT,
            "archivo_precios": ARCHIVO_PRECIOS_DEFAULT,
            "url_onedrive_alumnos": "",
            "archivo_alumnos": ARCHIVO_ALUMNOS_EJEMPLO,
            "frecuencia_sync_alumnos": "diaria",
            "ultima_fecha_emision": date.today().strftime("%d/%m/%Y"),
        }, f, indent=2, ensure_ascii=False)

    print(f"Listo: {len(nombres)} alumnos y recibos de ejemplo en {DIR_ACTUAL}")
    print("Ahora: python reemplazo_cuaderno.py")


if __name__ == "__main__":
    main()
