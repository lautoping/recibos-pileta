import os
import re
import shutil
import tempfile
import time
from copy import copy
from datetime import datetime, date
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import COLUMNAS_RECIBOS
from utils import quitar_tildes, parsear_fecha, formatear_moneda, limpiar_monto, sumar_un_mes


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
                            inscripcion_val = float(limpiar_monto(fila.iloc[c_idx + 1]))
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
                            carnet_val = float(limpiar_monto(fila.iloc[c_idx + 1]))
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
                            limpio = limpiar_monto(fila.iloc[col_i])
                            precios_default["cuotas"][clases_sem][desc_key] = float(limpio)
                        except ValueError:
                            pass

    precios_default["inscripcion"] = inscripcion_val
    precios_default["carnet"] = carnet_val
    precios_default["vigencia"] = vigencia_txt
    return precios_default


# ==========================================================
# HOJAS POR MES, GUARDADO SEGURO Y COPIAS DE SEGURIDAD
# ==========================================================
# Cada mes tiene su propia hoja dentro de recibos.xlsx ("Octubre 2026", "Noviembre 2026", ...).
# Un recibo va a la hoja del mes de su Fecha de Emisión. Las deudas se buscan en TODAS las hojas,
# así que una deuda del mes pasado sigue apareciendo hasta que se cobra.

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
CARPETA_COPIAS = "copias_seguridad"
COPIAS_A_CONSERVAR = 60

COL_NRO, COL_FECHA, COL_NOMBRE, COL_ACTIVIDAD = 1, 2, 3, 4
COL_METODO, COL_TOTAL, COL_PAGADO, COL_SALDO, COL_ESTADO, COL_OBS = 10, 14, 15, 16, 17, 18

ACTIVIDADES_NO_RECIBO = ("Cobro de Saldo",)

_BORDE = Border(
    bottom=Side(style="thin", color="E2E8F0"),
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
)
_FILL_HEADER = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
_FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
_FONT_NORMAL = Font(name="Segoe UI", size=10)
_ANCHOS = [12, 14, 28, 14, 24, 16, 16, 12, 16, 16, 13, 14, 16, 14, 15, 16, 14, 30]


def nombre_hoja_mes(fecha):
    """'Noviembre 2026' para una fecha (date o texto DD/MM/AAAA). None si no se puede leer."""
    if isinstance(fecha, (datetime, date)):
        dt = fecha
    else:
        dt = parsear_fecha(fecha) if fecha else None
    if not dt:
        return None
    return f"{MESES[dt.month - 1]} {dt.year}"


def _clave_hoja(titulo):
    """(año, mes) de una hoja mensual, o None si la hoja no es de un mes."""
    partes = str(titulo).strip().split()
    if len(partes) != 2 or not partes[1].isdigit():
        return None
    meses_norm = [quitar_tildes(m) for m in MESES]
    mes = quitar_tildes(partes[0])
    if mes not in meses_norm:
        return None
    return int(partes[1]), meses_norm.index(mes) + 1


def _hojas_ordenadas(wb):
    """Hojas mensuales de la más vieja a la más nueva."""
    hojas = [ws for ws in wb.worksheets if _clave_hoja(ws.title)]
    return sorted(hojas, key=lambda ws: _clave_hoja(ws.title))


def _preparar_hoja(ws):
    for col_idx, col_name in enumerate(COLUMNAS_RECIBOS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = _FILL_HEADER
        cell.font = _FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for idx, ancho in enumerate(_ANCHOS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = ancho
    ws.freeze_panes = "A2"


def _obtener_o_crear_hoja(wb, titulo):
    if titulo in wb.sheetnames:
        return wb[titulo], False
    ws = wb.create_sheet(titulo)
    _preparar_hoja(ws)
    # Ubicar las hojas en orden cronológico (las que no son de un mes quedan al final)
    wb._sheets.sort(key=lambda w: (0, _clave_hoja(w.title)) if _clave_hoja(w.title) else (1, (0, 0)))
    return ws, True


def _fila_vacia(ws, r):
    return all(ws.cell(row=r, column=c).value in (None, "") for c in range(1, len(COLUMNAS_RECIBOS) + 1))


def _ultima_fila_con_datos(ws):
    r = ws.max_row
    while r > 1 and _fila_vacia(ws, r):
        r -= 1
    return r


def _copia_de_seguridad(ruta):
    """Guarda una copia del archivo tal como estaba, una vez por día (antes del primer cambio del día)."""
    if not os.path.exists(ruta) or os.path.getsize(ruta) == 0:
        return
    carpeta = os.path.join(os.path.dirname(os.path.abspath(ruta)), CARPETA_COPIAS)
    base = os.path.splitext(os.path.basename(ruta))[0]
    destino = os.path.join(carpeta, f"{base}_{date.today().strftime('%Y-%m-%d')}.xlsx")
    if os.path.exists(destino):
        return
    try:
        os.makedirs(carpeta, exist_ok=True)
        shutil.copy2(ruta, destino)
        copias = sorted(f for f in os.listdir(carpeta) if f.startswith(base + "_") and f.endswith(".xlsx"))
        for viejo in copias[:-COPIAS_A_CONSERVAR]:
            try:
                os.remove(os.path.join(carpeta, viejo))
            except OSError:
                pass
    except Exception as e:
        print(f"Aviso: no se pudo crear la copia de seguridad: {e}")


def _guardar_wb(wb, ruta):
    """
    Guarda sin riesgo de dejar el Excel a medio escribir: primero en un archivo temporal
    y después lo reemplaza de una sola vez. Si recibos.xlsx está abierto en Excel, lanza PermissionError.
    """
    _copia_de_seguridad(ruta)
    carpeta = os.path.dirname(os.path.abspath(ruta))
    fd, tmp = tempfile.mkstemp(prefix="~guardando_", suffix=".xlsx", dir=carpeta)
    os.close(fd)
    try:
        wb.save(tmp)
        # OneDrive puede tener el archivo tomado un instante mientras sincroniza: reintentar un poco
        for intento in range(4):
            try:
                os.replace(tmp, ruta)
                break
            except PermissionError:
                if intento == 3:
                    raise
                time.sleep(0.5)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _migrar_columnas_antiguas(ws):
    """Recibos viejos sin las columnas de pago: Monto Pagado = Total, Saldo 0, Pagado."""
    headers_actuales = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if "Monto Pagado" in headers_actuales:
        return False
    for r in range(2, ws.max_row + 1):
        total_val = ws.cell(row=r, column=14).value or 0
        obs_val = ws.cell(row=r, column=15).value
        ws.cell(row=r, column=15, value=total_val)
        ws.cell(row=r, column=16, value=0)
        ws.cell(row=r, column=17, value="Pagado")
        ws.cell(row=r, column=18, value=obs_val)
    _preparar_hoja(ws)
    for r in range(2, ws.max_row + 1):
        _formatear_fila(ws, r, "Pagado")
    return True


def _separar_hoja_por_mes(wb, ws):
    """
    Pasa una hoja vieja (ej. 'Recibos') a hojas mensuales según la Fecha de Emisión.
    Las filas sin fecha quedan en el mes de la fila anterior.
    Si todo es de un solo mes, solo renombra la hoja (conserva todo el formato).
    """
    ultima = _ultima_fila_con_datos(ws)
    meses_fila = []
    mes_actual = None
    for r in range(2, ultima + 1):
        m = nombre_hoja_mes(ws.cell(row=r, column=COL_FECHA).value)
        if m:
            mes_actual = m
        meses_fila.append(mes_actual)
    primer_mes = next((m for m in meses_fila if m), None) or nombre_hoja_mes(date.today())
    meses_fila = [m or primer_mes for m in meses_fila]
    distintos = list(dict.fromkeys(meses_fila))

    if len(distintos) <= 1 and primer_mes not in wb.sheetnames:
        ws.title = primer_mes
        ws.freeze_panes = "A2"
        return

    for r, mes in zip(range(2, ultima + 1), meses_fila):
        destino, _ = _obtener_o_crear_hoja(wb, mes)
        nueva = _ultima_fila_con_datos(destino) + 1
        for c in range(1, len(COLUMNAS_RECIBOS) + 1):
            origen = ws.cell(row=r, column=c)
            celda = destino.cell(row=nueva, column=c, value=origen.value)
            if origen.has_style:
                celda._style = copy(origen._style)
    wb.remove(ws)


def asegurar_excel_recibos(ruta_recibos, fecha_referencia=None):
    """
    Crea o actualiza recibos.xlsx:
    - Una hoja por mes, con encabezados y formato.
    - Pasa la hoja vieja 'Recibos' (o cualquier hoja que no sea de un mes) a hojas mensuales.
    - Crea la hoja del mes actual si todavía no existe (el cambio de mes es automático).
    """
    mes_actual = nombre_hoja_mes(fecha_referencia or date.today())

    if not os.path.exists(ruta_recibos) or os.path.getsize(ruta_recibos) == 0:
        wb = Workbook()
        ws = wb.active
        ws.title = mes_actual
        _preparar_hoja(ws)
        _guardar_wb(wb, ruta_recibos)
        return

    try:
        wb = load_workbook(ruta_recibos)
        cambio = False
        for ws in list(wb.worksheets):
            if _clave_hoja(ws.title):
                cambio = _migrar_columnas_antiguas(ws) or cambio
                continue
            encabezado = str(ws.cell(row=1, column=1).value or "").strip()
            if encabezado == "Nro Recibo":
                cambio = _migrar_columnas_antiguas(ws) or cambio
                _separar_hoja_por_mes(wb, ws)
                cambio = True

        if mes_actual not in wb.sheetnames:
            _obtener_o_crear_hoja(wb, mes_actual)
            cambio = True

        if cambio:
            hojas = _hojas_ordenadas(wb)
            if hojas:
                wb.active = wb.worksheets.index(hojas[-1])
            _guardar_wb(wb, ruta_recibos)
    except PermissionError:
        raise
    except Exception as e:
        print(f"Aviso preparando el Excel de recibos: {e}")


def _leer_filas(wb):
    """Devuelve todas las filas de todas las hojas mensuales en orden cronológico, como dicts."""
    filas = []
    for ws in _hojas_ordenadas(wb):
        for r in range(2, _ultima_fila_con_datos(ws) + 1):
            valores = [ws.cell(row=r, column=c).value for c in range(1, len(COLUMNAS_RECIBOS) + 1)]
            if all(v in (None, "") for v in valores):
                continue
            d = {col: ("" if v is None else v) for col, v in zip(COLUMNAS_RECIBOS, valores)}
            d["_hoja"] = ws.title
            d["_fila"] = r
            filas.append(d)
    return filas


def _a_numero(valor):
    try:
        if valor is None or str(valor).strip() == "":
            return 0.0
        return float(valor)
    except (ValueError, TypeError):
        try:
            return float(limpiar_monto(valor))
        except ValueError:
            return 0.0


def _nro_normalizado(valor):
    """'R12', 12, 12.0 y ' 12 ' -> '12'."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip().lstrip("Rr").strip()


def _fecha_normalizada(valor):
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%Y")
    dt = parsear_fecha(valor) if valor not in (None, "") else None
    return dt.strftime("%d/%m/%Y") if dt else str(valor or "").strip()


def obtener_estado_recibos(ruta_recibos):
    """
    Devuelve:
    - proximo_recibo (int)
    - proximo_carnet (int)
    - ultima_fecha_emision (str DD/MM/YYYY o None)
    - ultimos_registros (lista de dicts para la tabla, el más nuevo primero)
    """
    asegurar_excel_recibos(ruta_recibos)
    try:
        wb = load_workbook(ruta_recibos, data_only=True)
    except Exception:
        return 1, 1, None, []

    filas = _leer_filas(wb)
    if not filas:
        return 1, 1, None, []

    # Se recorre desde el final para respetar cuando se termina un talonario y se reinicia desde 1
    proximo_recibo = 1
    for f in reversed(filas):
        act_val = str(f.get("Actividad", ""))
        if act_val in ACTIVIDADES_NO_RECIBO or act_val.startswith("Carnet"):
            continue
        m = re.match(r"^\s*[Rr]?\s*(\d+)", _nro_normalizado(f.get("Nro Recibo")))
        if m:
            proximo_recibo = int(m.group(1)) + 1
            break

    proximo_carnet = 1
    for f in reversed(filas):
        val_car = f.get("Nro Carnet")
        if val_car not in (None, ""):
            m = re.search(r"\b(\d+)\b", _nro_normalizado(val_car))
            if m:
                proximo_carnet = int(m.group(1)) + 1
                break

    ultima_fecha_emision = None
    for f in reversed(filas):
        if f.get("Fecha Emisión") not in (None, ""):
            ultima_fecha_emision = _fecha_normalizada(f["Fecha Emisión"])
            break

    ultimos = filas[-35:]
    ultimos.reverse()
    return proximo_recibo, proximo_carnet, ultima_fecha_emision, ultimos


def _formatear_fila(ws, fila, estado):
    """Formato de una fila; Saldo y Estado en amarillo si es 'Parcial' y en rojo si es 'Debe'."""
    fill_alerta = None
    font_alerta = None
    if estado == "Parcial":
        fill_alerta = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
        font_alerta = Font(name="Segoe UI", size=10, bold=True, color="854D0E")
    elif estado == "Debe":
        fill_alerta = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
        font_alerta = Font(name="Segoe UI", size=10, bold=True, color="991B1B")

    for col_idx in range(1, len(COLUMNAS_RECIBOS) + 1):
        c = ws.cell(row=fila, column=col_idx)
        c.font = _FONT_NORMAL
        c.border = _BORDE
        if col_idx in (1, 2, 6, 7, 8, 11, 17):
            c.alignment = Alignment(horizontal="center")
        elif col_idx in (12, 13, 14, 15, 16):
            if c.value != "" and c.value is not None:
                c.number_format = "$#,##0"
            c.alignment = Alignment(horizontal="right")
        if col_idx in (COL_SALDO, COL_ESTADO):
            if fill_alerta:
                c.fill = fill_alerta
                c.font = font_alerta
            else:
                c.fill = PatternFill(fill_type=None)


def _agregar_fila(wb, fila_datos):
    """Agrega la fila en la hoja del mes de su Fecha de Emisión (o del mes actual si no tiene)."""
    mes = nombre_hoja_mes(fila_datos.get("Fecha Emisión")) or nombre_hoja_mes(date.today())
    ws, _ = _obtener_o_crear_hoja(wb, mes)
    nueva = _ultima_fila_con_datos(ws) + 1
    for col_idx, col in enumerate(COLUMNAS_RECIBOS, start=1):
        valor = fila_datos.get(col, "")
        ws.cell(row=nueva, column=col_idx, value=None if valor == "" else valor)
    _formatear_fila(ws, nueva, fila_datos.get("Estado Pago", "Pagado"))
    return ws, nueva


def guardar_filas_recibo(ruta_recibos, filas_datos):
    """Guarda todas las filas de un recibo (ej. grupo familiar) en un solo guardado: o se guardan todas o ninguna."""
    asegurar_excel_recibos(ruta_recibos)
    wb = load_workbook(ruta_recibos)
    for fila in filas_datos:
        _agregar_fila(wb, fila)
    _guardar_wb(wb, ruta_recibos)


def guardar_fila_recibo(ruta_recibos, fila_datos):
    """Agrega una fila al final de la hoja del mes correspondiente manteniendo el formato."""
    guardar_filas_recibo(ruta_recibos, [fila_datos])


def _filas_del_recibo(ws, nro_recibo, fecha_emision, fila_ref=None):
    """
    Filas de una hoja con ese N° de recibo Y esa fecha de emisión (un mismo recibo, aunque sea grupal).
    Con fila_ref, solo el bloque de filas seguidas que la contiene (las personas de un recibo grupal
    se guardan juntas), así no se mezcla con otro recibo cargado con el mismo número el mismo día.
    """
    nro = _nro_normalizado(nro_recibo)
    fecha = _fecha_normalizada(fecha_emision)

    def coincide(r):
        return (
            _nro_normalizado(ws.cell(row=r, column=COL_NRO).value) == nro
            and _fecha_normalizada(ws.cell(row=r, column=COL_FECHA).value) == fecha
        )

    ultima = _ultima_fila_con_datos(ws)
    if fila_ref:
        if not (2 <= fila_ref <= ultima) or not coincide(fila_ref):
            return []
        ini = fila_ref
        while ini - 1 >= 2 and coincide(ini - 1):
            ini -= 1
        fin = fila_ref
        while fin + 1 <= ultima and coincide(fin + 1):
            fin += 1
        return list(range(ini, fin + 1))
    return [r for r in range(2, ultima + 1) if coincide(r)]


def buscar_recibo(ruta_recibos, nro_recibo, fecha_emision, fila_ref=None):
    """
    Filas de un recibo identificado por N° + fecha de emisión (así no se confunde con
    el mismo número de otro talonario). Devuelve lista de dicts con 'Nombre', 'Actividad', etc.
    """
    if not os.path.exists(ruta_recibos):
        return []
    mes = nombre_hoja_mes(fecha_emision)
    wb = load_workbook(ruta_recibos, data_only=True)
    if not mes or mes not in wb.sheetnames:
        return []
    ws = wb[mes]
    res = []
    for r in _filas_del_recibo(ws, nro_recibo, fecha_emision, fila_ref):
        res.append({col: ws.cell(row=r, column=i).value for i, col in enumerate(COLUMNAS_RECIBOS, start=1)})
    return res


def recibos_con_mismo_numero_en_el_mes(ruta_recibos, nro_recibo, fecha_emision):
    """Recibos (no cobros de saldo ni carnets) con ese número en la hoja del mes de la fecha dada."""
    if not os.path.exists(ruta_recibos):
        return []
    mes = nombre_hoja_mes(fecha_emision)
    wb = load_workbook(ruta_recibos, data_only=True)
    if not mes or mes not in wb.sheetnames:
        return []
    ws = wb[mes]
    nro = _nro_normalizado(nro_recibo)
    res = []
    for r in range(2, _ultima_fila_con_datos(ws) + 1):
        act = str(ws.cell(row=r, column=COL_ACTIVIDAD).value or "")
        if act in ACTIVIDADES_NO_RECIBO or act.startswith("Carnet"):
            continue
        if _nro_normalizado(ws.cell(row=r, column=COL_NRO).value) == nro:
            res.append({
                "fecha": _fecha_normalizada(ws.cell(row=r, column=COL_FECHA).value),
                "nombre": str(ws.cell(row=r, column=COL_NOMBRE).value or ""),
            })
    return res


def eliminar_recibo(ruta_recibos, nro_recibo, fecha_emision, fila_ref=None):
    """
    Elimina las filas del recibo con ese N° y esa fecha de emisión (todas las personas si es grupal).
    No toca recibos de otros talonarios con el mismo número. Devuelve cuántas filas borró.
    """
    mes = nombre_hoja_mes(fecha_emision)
    wb = load_workbook(ruta_recibos)
    if not mes or mes not in wb.sheetnames:
        return 0
    ws = wb[mes]
    filas = _filas_del_recibo(ws, nro_recibo, fecha_emision, fila_ref)
    for r in sorted(filas, reverse=True):
        ws.delete_rows(r, 1)
    if filas:
        _guardar_wb(wb, ruta_recibos)
    return len(filas)


def eliminar_recibo_por_numero(ruta_recibos, nro_recibo, fecha_emision=None):
    """Compatibilidad: ahora siempre hace falta la fecha para no borrar recibos de otro talonario."""
    if not fecha_emision:
        raise ValueError("Para eliminar un recibo hace falta su fecha de emisión.")
    return eliminar_recibo(ruta_recibos, nro_recibo, fecha_emision) > 0


def resumen_recibos(ruta_recibos):
    """Lista liviana de todos los recibos (todas las hojas) para buscar recibos anteriores de un grupo familiar."""
    if not os.path.exists(ruta_recibos):
        return []
    try:
        wb = load_workbook(ruta_recibos, read_only=True, data_only=True)
    except Exception:
        return []
    res = []
    hojas = sorted((w for w in wb.worksheets if _clave_hoja(w.title)), key=lambda w: _clave_hoja(w.title))
    for ws in hojas:
        for fila in ws.iter_rows(min_row=2, max_col=len(COLUMNAS_RECIBOS), values_only=True):
            if not fila or fila[COL_NOMBRE - 1] in (None, ""):
                continue
            res.append({
                "nro": _nro_normalizado(fila[COL_NRO - 1]),
                "fecha": _fecha_normalizada(fila[COL_FECHA - 1]),
                "nombre": str(fila[COL_NOMBRE - 1]).strip(),
                "actividad": str(fila[COL_ACTIVIDAD - 1] or ""),
            })
    wb.close()
    return res


def buscar_recibo_previo_del_grupo(resumen, otros_miembros, fecha_emision):
    """
    Busca el recibo más reciente de otro integrante del grupo familiar que esté vigente
    para la fecha de emisión dada (emitido como mucho un mes antes). None si no hay.
    """
    dt_actual = parsear_fecha(fecha_emision) if fecha_emision else None
    if not dt_actual or not otros_miembros:
        return None
    buscados = {quitar_tildes(m) for m in otros_miembros}
    mejor = None
    for r in resumen:
        if r["actividad"] in ACTIVIDADES_NO_RECIBO or r["actividad"].startswith("Carnet"):
            continue
        if r["actividad"] in ("Solo Inscripción", "Clase Adaptación (CdA)"):
            continue
        if quitar_tildes(r["nombre"]) not in buscados or not r["nro"]:
            continue
        dt = parsear_fecha(r["fecha"])
        if not dt or dt > dt_actual:
            continue
        if dt_actual > sumar_un_mes(dt):
            continue
        if mejor is None or dt >= mejor["_dt"]:
            mejor = dict(r, _dt=dt)
    if mejor:
        mejor.pop("_dt", None)
    return mejor


def obtener_recibos_con_deuda(ruta_recibos):
    """
    Recibos con saldo pendiente > 0 en TODAS las hojas (incluye las deudas de meses anteriores).
    Las filas de 'Cobro de Saldo' son solo el registro de un pago y nunca cuentan como deuda.
    El más reciente primero.
    """
    asegurar_excel_recibos(ruta_recibos)
    if not os.path.exists(ruta_recibos):
        return []

    wb = load_workbook(ruta_recibos, data_only=True)
    recibos_deuda = []
    for ws in _hojas_ordenadas(wb):
        for r in range(2, _ultima_fila_con_datos(ws) + 1):
            if str(ws.cell(row=r, column=COL_ACTIVIDAD).value or "") in ACTIVIDADES_NO_RECIBO:
                continue
            saldo_num = _a_numero(ws.cell(row=r, column=COL_SALDO).value)
            if saldo_num <= 0:
                continue
            recibos_deuda.append({
                "hoja": ws.title,
                "fila_excel": r,
                "nro_recibo": ws.cell(row=r, column=1).value,
                "fecha_emision": _fecha_normalizada(ws.cell(row=r, column=2).value),
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
                "estado_pago": str(ws.cell(row=r, column=17).value or "").strip(),
                "observaciones": str(ws.cell(row=r, column=18).value or ""),
            })

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
    hoja=None,
):
    """
    1. Actualiza la fila del recibo con deuda (ubicada por hoja + fila; si no se pasan,
       se busca la deuda más reciente con ese N° y nombre):
       - Mantiene intacto el Monto Pagado del día de emisión.
       - Descuenta el Saldo Pendiente (o lo deja en 0 si se salda en su totalidad).
       - Si el saldo queda en 0, cambia el estado a 'Pagado' y quita el color en las filas de ESE recibo
         (mismo N° y misma fecha de emisión; nunca toca otro talonario).
    2. Agrega una fila 'Cobro de Saldo' en la hoja del mes del pago (saldo 0: es solo el registro del cobro).
    Todo se guarda de una sola vez.
    """
    asegurar_excel_recibos(ruta_recibos)
    wb = load_workbook(ruta_recibos)
    monto_pago_float = float(monto_pago)

    ws = None
    fila = None
    if hoja and hoja in wb.sheetnames and fila_excel_especifica:
        ws = wb[hoja]
        fila = fila_excel_especifica
    else:
        nro_buscado = _nro_normalizado(nro_recibo)
        for ws_i in reversed(_hojas_ordenadas(wb)):
            for r in range(_ultima_fila_con_datos(ws_i), 1, -1):
                if str(ws_i.cell(row=r, column=COL_ACTIVIDAD).value or "") in ACTIVIDADES_NO_RECIBO:
                    continue
                if _nro_normalizado(ws_i.cell(row=r, column=COL_NRO).value) != nro_buscado:
                    continue
                if nombre_alumno and str(ws_i.cell(row=r, column=COL_NOMBRE).value or "").strip() != str(nombre_alumno).strip():
                    continue
                if _a_numero(ws_i.cell(row=r, column=COL_SALDO).value) > 0:
                    ws, fila = ws_i, r
                    break
            if ws:
                break
    if ws is None or not (2 <= fila <= ws.max_row):
        raise ValueError(f"No se encontró el recibo {nro_recibo} con saldo pendiente.")

    fecha_orig = _fecha_normalizada(ws.cell(row=fila, column=COL_FECHA).value)
    venc_orig = str(ws.cell(row=fila, column=7).value or "")
    filas_recibo = _filas_del_recibo(ws, ws.cell(row=fila, column=COL_NRO).value, fecha_orig, fila) or [fila]

    total_orig = max([_a_numero(ws.cell(row=r, column=COL_TOTAL).value) for r in filas_recibo] + [0.0])

    saldo_actual = _a_numero(ws.cell(row=fila, column=COL_SALDO).value)
    a_descontar = min(monto_pago_float, saldo_actual)
    nuevo_saldo = 0.0 if saldar_total else max(0.0, saldo_actual - a_descontar)

    ws.cell(row=fila, column=COL_SALDO, value=nuevo_saldo)
    if str(ws.cell(row=fila, column=COL_METODO).value or "").strip().lower() in ("no figura", ""):
        ws.cell(row=fila, column=COL_METODO, value=metodo_pago)

    if nuevo_saldo <= 0:
        nota = f"[Saldo cancelado {fecha_pago} ({metodo_pago})]"
        # Recibo grupal: todas las filas de ESTE recibo quedan en Pagado y sin color
        for r in filas_recibo:
            ws.cell(row=r, column=COL_ESTADO, value="Pagado")
            if _a_numero(ws.cell(row=r, column=COL_SALDO).value) > 0:
                ws.cell(row=r, column=COL_SALDO, value=0)
            for c in (COL_SALDO, COL_ESTADO):
                ws.cell(row=r, column=c).fill = PatternFill(fill_type=None)
                ws.cell(row=r, column=c).font = _FONT_NORMAL
    else:
        nota = f"[Abonó {formatear_moneda(a_descontar)} el {fecha_pago} ({metodo_pago}) - Resta {formatear_moneda(nuevo_saldo)}]"
        fill_amarillo = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
        font_amarillo = Font(name="Segoe UI", size=10, bold=True, color="854D0E")
        ws.cell(row=fila, column=COL_ESTADO, value="Parcial")
        for c in (COL_SALDO, COL_ESTADO):
            ws.cell(row=fila, column=c).fill = fill_amarillo
            ws.cell(row=fila, column=c).font = font_amarillo

    obs_prev = str(ws.cell(row=fila, column=COL_OBS).value or "").strip()
    ws.cell(row=fila, column=COL_OBS, value=f"{obs_prev} {nota}".strip())

    # 2. Fila de registro del cobro (en la hoja del mes del pago)
    nro_txt = _nro_normalizado(ws.cell(row=fila, column=COL_NRO).value) or _nro_normalizado(nro_recibo)
    det_act = f"Abono saldo R{nro_txt}"
    if fecha_orig:
        det_act += f" del {fecha_orig}"

    obs_pago = f"Abonó {formatear_moneda(monto_pago_float)} en {metodo_pago}, R{nro_txt}"
    if fecha_orig:
        obs_pago += f" del {fecha_orig}"
    if nuevo_saldo > 0:
        obs_pago += f" - Resta {formatear_moneda(nuevo_saldo)} (queda en el recibo original)"
    if observaciones_extra:
        obs_pago += f" | {observaciones_extra}"

    fila_registro = {
        "Nro Recibo": f"R{nro_txt}",
        "Fecha Emisión": fecha_pago,
        "Nombre": nombre_alumno or str(ws.cell(row=fila, column=COL_NOMBRE).value or ""),
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
        "Saldo Pendiente": 0,
        "Estado Pago": "Pagado",
        "Observaciones": obs_pago,
    }
    _agregar_fila(wb, fila_registro)
    _guardar_wb(wb, ruta_recibos)
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
