import re
import calendar
import unicodedata
from datetime import datetime, date


def quitar_tildes(texto):
    """Elimina acentos y normaliza a minúsculas para búsquedas flexibles."""
    if not texto:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    ).lower().strip()


def formatear_nombre_propio(texto):
    """
    Normaliza espacios y pone en mayúscula la primera letra de cada nombre y apellido
    (ej. 'juan   perez' -> 'Juan Perez', 'gómez, maría' -> 'Gómez, María').
    """
    if not texto:
        return ""
    limpio = " ".join(str(texto).split())
    return re.sub(r"[^\W\d_]+", lambda m: m.group(0).capitalize(), limpio)


def sumar_un_mes(fecha_dt):
    """Suma exactamente 1 mes calendario a un objeto date/datetime."""
    anio = fecha_dt.year
    mes = fecha_dt.month + 1
    if mes > 12:
        mes = 1
        anio += 1
    ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
    dia = min(fecha_dt.day, ultimo_dia_mes)
    return date(anio, mes, dia)


def parsear_fecha(texto_fecha):
    """Intenta convertir un string DD/MM/YYYY o DD/MM/YY a date."""
    texto = str(texto_fecha).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def formatear_moneda(valor):
    """Formatea un número flotante o entero como moneda argentina (ej: $73.000)."""
    try:
        num = float(valor)
        if num.is_integer():
            return f"${int(num):,}".replace(",", ".")
        return f"${num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "$0"
