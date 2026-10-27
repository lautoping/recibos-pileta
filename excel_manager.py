import os
import re
from datetime import datetime, date
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import COLUMNAS_RECIBOS
from utils import quitar_tildes, parsear_fecha, formatear_moneda


def crear_excel_precios_por_defecto(ruta_precios):
    """Crea el archivo precios.xlsx con la estructura estándar."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Precios"

    ws["A1"] = "A partir del 14/9"
    ws["A2"] = "Inscripción $15000."

    encabezados = ["Precios", "Precio", "Pago Efectivo", "Gpo familiar", "Jubilados"]
    ws.append(encabezados)

    filas = [
        ["4 clases/mes", 41000, 37000, 33000, 37000],
        ["8 clases/mes", 73000, 66000, 58000, 66000],
        ["12 clases/mes", 103000, 94000, 82000, 94000],
    ]
    for f in filas:
        ws.append(f)

    ws.append(["Descuentos no acumulables"])
    ws.append(["Reposición carnet $1000."])

    borde_fino = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )

    for row in ws.iter_rows(min_row=1, max_row=8, min_col=1, max_col=5):
        for cell in row:
            cell.font = Font(name="Calibri", size=11)
            cell.border = borde_fino
            if 3 <= cell.row <= 6:
                cell.alignment = Alignment(horizontal="center")

    anchos = [24, 14, 16, 16, 14]
    for idx, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = ancho

    wb.save(ruta_precios)


def cargar_tabla_precios(ruta_precios):
    """
    Lee el Excel de precios:
    - Detecta el valor de 'Inscripción'
    - Detecta precios según clases semanales y tipo de descuento.
    - Detecta el valor de 'Reposición carnet'
    """
    precios_default = {
        "inscripcion": 15000.0,
        "carnet": 1000.0,
        "vigencia": "",
        "cuotas": {
            1: {"Ninguno": 41000.0, "Efectivo": 37000.0, "Grupo familiar": 33000.0, "Jubilado": 37000.0},
            2: {"Ninguno": 73000.0, "Efectivo": 66000.0, "Grupo familiar": 58000.0, "Jubilado": 66000.0},
            3: {"Ninguno": 103000.0, "Efectivo": 94000.0, "Grupo familiar": 82000.0, "Jubilado": 94000.0},
        },
    }

    if not os.path.exists(ruta_precios):
        crear_excel_precios_por_defecto(ruta_precios)

    try:
        df_raw = pd.read_excel(ruta_precios, header=None)
    except Exception:
        return precios_default

    inscripcion_val = precios_default["inscripcion"]
    carnet_val = precios_default["carnet"]
    vigencia_txt = ""
    fila_header = None
    mapa_columnas = {}

    for r_idx, fila in df_raw.iterrows():
        for c_idx, val in enumerate(fila):
            if pd.isna(val):
                continue
            texto = str(val).strip()
            texto_norm = quitar_tildes(texto)

            if "a partir de" in texto_norm:
                vigencia_txt = texto

            if "inscrip" in texto_norm:
                numeros = re.findall(r"\d+(?:[\.,]\d+)?", texto.replace(".", ""))
                if numeros:
                    try:
                        inscripcion_val = float(numeros[0].replace(",", "."))
                    except ValueError:
                        pass
                else:
                    if c_idx + 1 < len(fila) and not pd.isna(fila.iloc[c_idx + 1]):
                        try:
                            inscripcion_val = float(str(fila.iloc[c_idx + 1]).replace("$", "").replace(".", "").strip())
                        except ValueError:
                            pass

            if "carnet" in texto_norm or "reposic" in texto_norm:
                numeros = re.findall(r"\d+(?:[\.,]\d+)?", texto.replace(".", ""))
                if numeros:
                    try:
                        carnet_val = float(numeros[0].replace(",", "."))
                    except ValueError:
                        pass
                else:
                    if c_idx + 1 < len(fila) and not pd.isna(fila.iloc[c_idx + 1]):
                        try:
                            carnet_val = float(str(fila.iloc[c_idx + 1]).replace("$", "").replace(".", "").strip())
                        except ValueError:
                            pass

            if texto_norm in ("precio", "pago efectivo", "gpo familiar", "grupo familiar", "jubilados", "jubilado"):
                fila_header = r_idx

    if fila_header is not None:
        fila_enc = df_raw.iloc[fila_header]
        for c_idx, val in enumerate(fila_enc):
            if pd.isna(val):
                continue
            t = quitar_tildes(val)
            if t == "precio" or "lista" in t or "normal" in t:
                mapa_columnas["Ninguno"] = c_idx
            elif "efectivo" in t:
                mapa_columnas["Efectivo"] = c_idx
            elif "familiar" in t:
                mapa_columnas["Grupo familiar"] = c_idx
            elif "jubilad" in t:
                mapa_columnas["Jubilado"] = c_idx

        for r_idx in range(fila_header + 1, len(df_raw)):
            fila = df_raw.iloc[r_idx]
            primera_celda = str(fila.iloc[0]) if not pd.isna(fila.iloc[0]) else ""
            t_celda = quitar_tildes(primera_celda)

            clases_sem = None
            if re.search(r"\b(12|3)\s*clases?", t_celda) or t_celda.startswith("12"):
                clases_sem = 3
            elif re.search(r"\b(8|2)\s*clases?", t_celda) or t_celda.startswith("8"):
                clases_sem = 2
            elif re.search(r"\b(4|1)\s*clases?", t_celda) or t_celda.startswith("4"):
                clases_sem = 1

            if clases_sem in (1, 2, 3):
                for desc_key, col_i in mapa_columnas.items():
                    if col_i < len(fila) and not pd.isna(fila.iloc[col_i]):
                        try:
                            limpio = str(fila.iloc[col_i]).replace("$", "").replace(" ", "").strip()
                            precios_default["cuotas"][clases_sem][desc_key] = float(limpio)
                        except ValueError:
                            pass

    precios_default["inscripcion"] = inscripcion_val
    precios_default["carnet"] = carnet_val
    precios_default["vigencia"] = vigencia_txt
    return precios_default


def asegurar_excel_recibos(ruta_recibos):
    """Crea o actualiza recibos.xlsx con los encabezados y columnas correctos."""
    borde_fino = Border(
        bottom=Side(style="thin", color="E2E8F0"),
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
    )
    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    anchos = [12, 14, 28, 14, 24, 16, 16, 12, 16, 16, 13, 14, 16, 14, 15, 16, 14, 30]

    if not os.path.exists(ruta_recibos) or os.path.getsize(ruta_recibos) == 0:
        wb = Workbook()
        ws = wb.active
        ws.title = "Recibos"
        ws.append(COLUMNAS_RECIBOS)

        for col_idx in range(1, len(COLUMNAS_RECIBOS) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for idx, ancho in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = ancho

        wb.save(ruta_recibos)
        return

    # Si ya existe, verificar si requiere migrar columnas antiguas
    try:
        wb = load_workbook(ruta_recibos)
        ws = wb.active
        headers_actuales = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

        if "Monto Pagado" not in headers_actuales:
            for r in range(2, ws.max_row + 1):
                total_val = ws.cell(row=r, column=14).value or 0
                obs_val = ws.cell(row=r, column=15).value

                ws.cell(row=r, column=15, value=total_val)       # Monto Pagado = Total
                ws.cell(row=r, column=16, value=0)               # Saldo Pendiente = 0
                ws.cell(row=r, column=17, value="Pagado")        # Estado Pago = "Pagado"
                ws.cell(row=r, column=18, value=obs_val)         # Observaciones

            for col_idx, col_name in enumerate(COLUMNAS_RECIBOS, start=1):
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for idx, ancho in enumerate(anchos, start=1):
                ws.column_dimensions[get_column_letter(idx)].width = ancho

            for r in range(2, ws.max_row + 1):
                for c in range(1, len(COLUMNAS_RECIBOS) + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.font = Font(name="Segoe UI", size=10)
                    cell.border = borde_fino
                    if c in (1, 2, 6, 7, 8, 11, 17):
                        cell.alignment = Alignment(horizontal="center")
                    elif c in (12, 13, 14, 15, 16):
                        cell.number_format = "$#,##0"
                        cell.alignment = Alignment(horizontal="right")

            wb.save(ruta_recibos)
    except Exception as e:
        print(f"Aviso migrando Excel de recibos: {e}")


def obtener_estado_recibos(ruta_recibos):
    """
    Devuelve:
    - proximo_recibo (int)
    - proximo_carnet (int)
    - ultima_fecha_emision (str DD/MM/YYYY o None)
    - ultimos_registros (lista de dicts para la tabla)
    """
    asegurar_excel_recibos(ruta_recibos)
    try:
        df = pd.read_excel(ruta_recibos)
    except Exception:
        return 1, 1, None, []

    if df.empty:
        return 1, 1, None, []

    proximo_recibo = 1
    if "Nro Recibo" in df.columns:
        # Buscar el último recibo emitido recorriendo el archivo desde el final hacia el inicio
        # para respetar cuando se termina un talonario y se reinicia desde 1
        for idx in range(len(df) - 1, -1, -1):
            val_rec = df.at[idx, "Nro Recibo"]
            act_val = str(df.at[idx, "Actividad"]) if "Actividad" in df.columns and not pd.isna(df.at[idx, "Actividad"]) else ""
            if act_val == "Cobro de Saldo" or act_val.startswith("Carnet"):
                continue

            if not pd.isna(val_rec):
                m = re.match(r"^\s*[Rr]?\s*(\d+)", str(val_rec).strip())
                if m:
                    proximo_recibo = int(m.group(1)) + 1
                    break

    proximo_carnet = 1
    if "Nro Carnet" in df.columns:
        # Buscar el último carnet asignado recorriendo desde el final hacia el inicio
        # para respetar cuando se termina una caja/serie de carnets y se reinicia desde 1
        for idx in range(len(df) - 1, -1, -1):
            val_car = df.at[idx, "Nro Carnet"]
            if not pd.isna(val_car):
                m = re.search(r"\b(\d+)\b", str(val_car).strip())
                if m:
                    proximo_carnet = int(m.group(1)) + 1
                    break

    ultima_fecha_emision = None
    if "Fecha Emisión" in df.columns:
        fechas_no_nulas = df["Fecha Emisión"].dropna()
        if not fechas_no_nulas.empty:
            ult_val = fechas_no_nulas.iloc[-1]
            if isinstance(ult_val, (datetime, date, pd.Timestamp)):
                ultima_fecha_emision = ult_val.strftime("%d/%m/%Y")
            else:
                dt_parsed = parsear_fecha(str(ult_val))
                if dt_parsed:
                    ultima_fecha_emision = dt_parsed.strftime("%d/%m/%Y")
                else:
                    ultima_fecha_emision = str(ult_val).strip()

    ultimos = df.tail(35).fillna("").to_dict(orient="records")
    ultimos.reverse()
    return proximo_recibo, proximo_carnet, ultima_fecha_emision, ultimos


def guardar_fila_recibo(ruta_recibos, fila_datos):
    """
    Agrega una fila al final de recibos.xlsx manteniendo el formato.
    Si el estado de pago es 'Parcial', pinta las celdas de Saldo Pendiente y Estado en amarillo.
    Si el estado de pago es 'Debe' (debe todo), las pinta en rojo.
    """
    asegurar_excel_recibos(ruta_recibos)
    wb = load_workbook(ruta_recibos)
    ws = wb.active

    valores = [fila_datos.get(col, "") for col in COLUMNAS_RECIBOS]
    ws.append(valores)

    ultima_fila = ws.max_row
    borde_fila = Border(
        bottom=Side(style="thin", color="E2E8F0"),
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
    )

    estado = fila_datos.get("Estado Pago", "Pagado")

    fill_alerta = None
    font_alerta = None

    if estado == "Parcial":
        # Amarillo para pago parcial (a medias)
        fill_alerta = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
        font_alerta = Font(name="Segoe UI", size=10, bold=True, color="854D0E")
    elif estado == "Debe":
        # Rojo para debe todo
        fill_alerta = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
        font_alerta = Font(name="Segoe UI", size=10, bold=True, color="991B1B")

    for col_idx in range(1, len(COLUMNAS_RECIBOS) + 1):
        c = ws.cell(row=ultima_fila, column=col_idx)
        c.font = Font(name="Segoe UI", size=10)
        c.border = borde_fila
        if col_idx in (1, 2, 6, 7, 8, 11, 17):
            c.alignment = Alignment(horizontal="center")
        elif col_idx in (12, 13, 14, 15, 16):
            if c.value != "" and c.value is not None:
                c.number_format = "$#,##0"
            c.alignment = Alignment(horizontal="right")

        # Pintar recuadros de Saldo Pendiente (col 16) y Estado Pago (col 17)
        if fill_alerta and col_idx in (16, 17):
            c.fill = fill_alerta
            c.font = font_alerta

    wb.save(ruta_recibos)


def eliminar_recibo_por_numero(ruta_recibos, nro_recibo):
    """Elimina todas las filas cuyo Nro Recibo coincida con nro_recibo."""
    wb = load_workbook(ruta_recibos)
    ws = wb.active
    filas_a_borrar = []

    for row_idx in range(ws.max_row, 1, -1):
        val = ws.cell(row=row_idx, column=1).value
        if val is not None and str(val).strip() == str(nro_recibo).strip():
            filas_a_borrar.append(row_idx)

    if filas_a_borrar:
        for r_idx in filas_a_borrar:
            ws.delete_rows(r_idx, 1)
        wb.save(ruta_recibos)
        return True
    return False


def obtener_recibos_con_deuda(ruta_recibos):
    """
    Devuelve la lista de registros de recibos.xlsx que tienen saldo pendiente > 0
    o cuyo estado sea 'Parcial' o 'Debe'.
    """
    asegurar_excel_recibos(ruta_recibos)
    if not os.path.exists(ruta_recibos):
        return []

    wb = load_workbook(ruta_recibos, data_only=True)
    ws = wb.active

    recibos_deuda = []
    for r in range(2, ws.max_row + 1):
        saldo_val = ws.cell(row=r, column=16).value
        try:
            saldo_num = float(saldo_val) if saldo_val is not None else 0.0
        except ValueError:
            saldo_num = 0.0

        estado_val = str(ws.cell(row=r, column=17).value or "").strip()

        if saldo_num > 0:
            recibo_info = {
                "fila_excel": r,
                "nro_recibo": ws.cell(row=r, column=1).value,
                "fecha_emision": str(ws.cell(row=r, column=2).value or ""),
                "nombre": str(ws.cell(row=r, column=3).value or ""),
                "actividad": str(ws.cell(row=r, column=4).value or ""),
                "clases_sem": ws.cell(row=r, column=6).value,
                "fecha_vencimiento": str(ws.cell(row=r, column=7).value or ""),
                "inscripcion": str(ws.cell(row=r, column=8).value or ""),
                "descuento": str(ws.cell(row=r, column=9).value or ""),
                "metodo_pago": str(ws.cell(row=r, column=10).value or ""),
                "total": ws.cell(row=r, column=14).value or 0,
                "monto_pagado": ws.cell(row=r, column=15).value or 0,
                "saldo_pendiente": saldo_num,
                "estado_pago": estado_val,
                "observaciones": str(ws.cell(row=r, column=18).value or ""),
            }
            recibos_deuda.append(recibo_info)

    recibos_deuda.reverse()
    return recibos_deuda


def registrar_pago_pendiente(
    ruta_recibos,
    nro_recibo,
    nombre_alumno,
    fecha_pago,
    monto_pago,
    metodo_pago,
    observaciones_extra="",
    fila_excel_especifica=None,
    saldar_total=False,
):
    """
    1. Actualiza la/s fila/s original/es del recibo en Excel:
       - Mantiene intacto el Monto Pagado del día de emisión.
       - Descuenta o cancela el Saldo Pendiente (0 si se salda en su totalidad).
       - Si el saldo queda en 0, cambia el estado a 'Pagado' y QUITA el color rojo/amarillo.
    2. Agrega una nueva fila al final del Excel con el detalle del abono:
       Total = total original, Monto Pagado = monto abonado, Saldo = 0 (o restante).
    """
    asegurar_excel_recibos(ruta_recibos)
    wb = load_workbook(ruta_recibos)
    ws = wb.active

    fecha_orig = ""
    venc_orig = ""
    monto_pago_float = float(monto_pago)
    saldo_restante_a_imputar = monto_pago_float

    filas_a_procesar = []
    if fila_excel_especifica and 2 <= fila_excel_especifica <= ws.max_row:
        filas_a_procesar.append(fila_excel_especifica)
    else:
        for r in range(2, ws.max_row + 1):
            val_rec = ws.cell(row=r, column=1).value
            if val_rec is not None and str(val_rec).strip().lstrip("Rr") == str(nro_recibo).strip().lstrip("Rr"):
                if not nombre_alumno or str(ws.cell(row=r, column=3).value).strip() == str(nombre_alumno).strip():
                    filas_a_procesar.append(r)

    total_orig = 0.0
    val_rec_buscado = str(nro_recibo).strip().lstrip("Rr")
    for r2 in range(2, ws.max_row + 1):
        vr = str(ws.cell(row=r2, column=1).value or "").strip().lstrip("Rr")
        if vr == val_rec_buscado:
            tot_cell = ws.cell(row=r2, column=14).value
            try:
                if tot_cell is not None and str(tot_cell).strip() != "":
                    t_val = float(tot_cell)
                    if t_val > total_orig:
                        total_orig = t_val
            except ValueError:
                pass

    ultimo_nuevo_saldo = 0.0
    for r in filas_a_procesar:
        if not fecha_orig:
            fecha_orig = str(ws.cell(row=r, column=2).value or "")
        if not venc_orig:
            venc_orig = str(ws.cell(row=r, column=7).value or "")

        saldo_actual_val = ws.cell(row=r, column=16).value
        try:
            saldo_actual = float(saldo_actual_val) if saldo_actual_val is not None else 0.0
        except ValueError:
            saldo_actual = 0.0

        pagado_actual_val = ws.cell(row=r, column=15).value
        try:
            pagado_actual = float(pagado_actual_val) if pagado_actual_val is not None else 0.0
        except ValueError:
            pagado_actual = 0.0

        if saldo_actual > 0 and (saldo_restante_a_imputar > 0 or saldar_total):
            a_descontar = min(saldo_restante_a_imputar, saldo_actual)
            nuevo_saldo = 0.0 if saldar_total else (saldo_actual - a_descontar)
            ultimo_nuevo_saldo = nuevo_saldo
            saldo_restante_a_imputar = max(0.0, saldo_restante_a_imputar - a_descontar)

            # En la fila original, el valor que se pagó el día anterior sigue igual (NO se incrementa)
            # y se actualiza el saldo pendiente a 0 (o lo que reste)
            ws.cell(row=r, column=16, value=nuevo_saldo)

            # Si el método de pago original era "No figura", actualizarlo al método real abonado
            if str(ws.cell(row=r, column=10).value or "").strip().lower() in ("no figura", ""):
                ws.cell(row=r, column=10, value=metodo_pago)

            obs_prev = str(ws.cell(row=r, column=18).value or "").strip()

            c16 = ws.cell(row=r, column=16)
            c17 = ws.cell(row=r, column=17)

            if nuevo_saldo <= 0:
                ws.cell(row=r, column=17, value="Pagado")
                # Quitar el color rojo o amarillo
                c16.fill = PatternFill(fill_type=None)
                c16.font = Font(name="Segoe UI", size=10)
                c17.fill = PatternFill(fill_type=None)
                c17.font = Font(name="Segoe UI", size=10)
                nota = f"[Saldo cancelado {fecha_pago} ({metodo_pago})]"

                # Si es un recibo con múltiples filas (grupo familiar), asegurar que todas queden en Pagado sin color
                for r2 in range(2, ws.max_row + 1):
                    vr = str(ws.cell(row=r2, column=1).value or "").strip().lstrip("Rr")
                    if vr == val_rec_buscado:
                        ws.cell(row=r2, column=17, value="Pagado")
                        ws.cell(row=r2, column=16).fill = PatternFill(fill_type=None)
                        ws.cell(row=r2, column=16).font = Font(name="Segoe UI", size=10)
                        ws.cell(row=r2, column=17).fill = PatternFill(fill_type=None)
                        ws.cell(row=r2, column=17).font = Font(name="Segoe UI", size=10)
            else:
                ws.cell(row=r, column=17, value="Parcial")
                fill_amarillo = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
                font_amarillo = Font(name="Segoe UI", size=10, bold=True, color="854D0E")
                c16.fill = fill_amarillo
                c16.font = font_amarillo
                c17.fill = fill_amarillo
                c17.font = font_amarillo
                nota = f"[Abonó {formatear_moneda(a_descontar)} el {fecha_pago} ({metodo_pago}) - Resta {formatear_moneda(nuevo_saldo)}]"

            ws.cell(row=r, column=18, value=f"{obs_prev} {nota}".strip())

    wb.save(ruta_recibos)

    # 2. Agregar la fila de registro del pago pendiente
    det_act = f"Abono saldo R{nro_recibo}"
    if fecha_orig:
        det_act += f" del {fecha_orig}"

    obs_pago = f"Abonó {formatear_moneda(monto_pago_float)} en {metodo_pago}, R{nro_recibo}"
    if fecha_orig:
        obs_pago += f" del {fecha_orig}"
    if observaciones_extra:
        obs_pago += f" | {observaciones_extra}"

    fila_registro = {
        "Nro Recibo": f"R{nro_recibo}",
        "Fecha Emisión": fecha_pago,
        "Nombre": nombre_alumno,
        "Actividad": "Cobro de Saldo",
        "Detalle Actividad": det_act,
        "Clases por Semana": "",
        "Fecha Vencimiento": venc_orig,
        "Inscripción": "",
        "Descuento": "",
        "Método de Pago": metodo_pago,
        "Nro Carnet": "",
        "Monto Cuota": 0,
        "Monto Inscripción": 0,
        "Total": total_orig if total_orig > 0 else monto_pago_float,
        "Monto Pagado": monto_pago_float,
        "Saldo Pendiente": max(0.0, ultimo_nuevo_saldo),
        "Estado Pago": "Pagado" if ultimo_nuevo_saldo <= 0 else "Parcial",
        "Observaciones": obs_pago,
    }

    guardar_fila_recibo(ruta_recibos, fila_registro)
    return True


def registrar_carnet_sin_recibo(
    ruta_recibos,
    nombre_alumno,
    nro_carnet,
    recibo_asociado,
    fecha_emision,
    modo_cobro,  # "Cobrado", "Sin cargo", "Debe"
    precio_carnet=1000.0,
    metodo_pago="Efectivo",
    observaciones_extra="",
    fecha_recibo_asociado="",
):
    """
    Registra en recibos.xlsx un carnet emitido sin recibo de cuota o como reposición.
    Maneja si se cobró, fue sin cargo o quedó debiendo.
    """
    asegurar_excel_recibos(ruta_recibos)
    precio_float = float(precio_carnet)

    if modo_cobro == "Cobrado":
        total = precio_float
        pagado = precio_float
        saldo = 0
        estado = "Pagado"
        mp = metodo_pago
        obs_cobro = f"Cobrado {formatear_moneda(precio_float)} ({metodo_pago})"
    elif modo_cobro == "Sin cargo":
        total = 0
        pagado = 0
        saldo = 0
        estado = "Pagado"
        mp = "Sin cargo"
        obs_cobro = "Sin cargo (no se cobra)"
    else:  # "Debe"
        total = precio_float
        pagado = 0
        saldo = precio_float
        estado = "Debe"
        mp = "Debe"
        obs_cobro = f"Debe {formatear_moneda(precio_float)}"

    rec_txt = ""
    nro_rec_guardar = ""
    if recibo_asociado:
        rec_limpio = str(recibo_asociado).strip()
        num_solo = rec_limpio.lstrip("Rr").strip()
        if num_solo.isdigit():
            rec_txt = f"R{num_solo}"
        else:
            rec_txt = rec_limpio if rec_limpio.startswith(("R", "r")) else f"R{rec_limpio}"
        nro_rec_guardar = rec_txt

        if fecha_recibo_asociado:
            rec_txt += f" del {fecha_recibo_asociado.strip()}"

    if rec_txt:
        detalle = f"Carnet N°{nro_carnet}, {rec_txt}"
    else:
        detalle = f"Carnet N°{nro_carnet}"

    obs_carnet = f"Carnet N°{nro_carnet} para {nombre_alumno}"
    if rec_txt:
        obs_carnet += f", {rec_txt}"
    obs_carnet += f" - {obs_cobro}"
    if observaciones_extra:
        obs_carnet += f" | {observaciones_extra}"

    fila_carnet = {
        "Nro Recibo": nro_rec_guardar,
        "Fecha Emisión": fecha_emision,
        "Nombre": nombre_alumno,
        "Actividad": "Carnet",
        "Detalle Actividad": detalle,
        "Clases por Semana": "",
        "Fecha Vencimiento": "",
        "Inscripción": "",
        "Descuento": "Sin cargo" if modo_cobro == "Sin cargo" else "",
        "Método de Pago": mp,
        "Nro Carnet": nro_carnet,
        "Monto Cuota": 0,
        "Monto Inscripción": 0,
        "Total": total,
        "Monto Pagado": pagado,
        "Saldo Pendiente": saldo,
        "Estado Pago": estado,
        "Observaciones": obs_carnet,
    }

    guardar_fila_recibo(ruta_recibos, fila_carnet)
    return True
