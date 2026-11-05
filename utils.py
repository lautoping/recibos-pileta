import re
import calendar
import unicodedata
from datetime import datetime, date, timedelta


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


DIAS_SEMANA = {
    "lunes": 0, "martes": 1, "miercoles": 2, "jueves": 3, "viernes": 4, "sabado": 5, "domingo": 6,
}
NOMBRE_DIA_CORTO = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]


def dias_semana_de_texto(dias_clase):
    """
    'Lunes, Miércoles' -> {0, 2}. None si el texto está vacío o tiene algo que no es un día
    (ej. un horario que no se pudo interpretar).
    """
    partes = [quitar_tildes(p) for p in str(dias_clase or "").split(",") if p.strip()]
    if not partes or any(p not in DIAS_SEMANA for p in partes):
        return None
    return {DIAS_SEMANA[p] for p in partes}


def revisar_atraso(vto, fecha_pago, dias_clase):
    """
    Pago atrasado = después del vto hubo una clase del alumno antes del día en que paga.
    (Vence el lunes 4 y viene los miércoles: si paga el miércoles 6 está en término.
    Vence el lunes 4, viene el martes 5 y paga el jueves 7: atrasado.)
    El día del vto y el día del pago no cuentan como clase "sin pagar".
    Devuelve (estado, fecha_de_la_clase):
      'sin_fechas'  vto o fecha de pago vacíos/inválidos
      'en_termino'  paga hasta el vto, o no tuvo clase entre el vto y el pago
      'atrasado'    tuvo clase entre el vto y el pago (devuelve la primera)
      'sin_dias'    paga después del vto pero no se sabe qué días viene
    """
    dt_vto = vto if isinstance(vto, date) else parsear_fecha(vto) if vto else None
    dt_pago = fecha_pago if isinstance(fecha_pago, date) else parsear_fecha(fecha_pago) if fecha_pago else None
    if not dt_vto or not dt_pago:
        return "sin_fechas", None
    if dt_pago <= dt_vto:
        return "en_termino", None
    dias = dias_semana_de_texto(dias_clase)
    if dias is None:
        return "sin_dias", None
    d = dt_vto + timedelta(days=1)
    while d < dt_pago:
        if d.weekday() in dias:
            return "atrasado", d
        d += timedelta(days=1)
    return "en_termino", None


def formatear_moneda(valor):
    """Formatea un número flotante o entero como moneda argentina (ej: $73.000)."""
    try:
        num = float(valor)
        if num.is_integer():
            return f"${int(num):,}".replace(",", ".")
        return f"${num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "$0"


def limpiar_monto(texto):
    """
    Normaliza un monto escrito a mano a un string que float() entiende.
    Acepta '$73.000', '73000', '1.234,50', '1234,5' y '1234.50'
    (un punto seguido de exactamente 3 cifras se toma como separador de miles).
    """
    t = str(texto if texto is not None else "").replace("$", "").replace(" ", "").strip()
    if "," in t:
        return t.replace(".", "").replace(",", ".")
    if "." in t and re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):
        return t.replace(".", "")
    return t
