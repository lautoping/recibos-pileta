import os
import io
import re
import json
import base64
import urllib.request
import urllib.parse
from datetime import datetime
import pandas as pd

from config import ARCHIVO_CACHE_ALUMNOS
from utils import quitar_tildes, formatear_nombre_propio


def extraer_dias_clase(texto):
    """
    Extrae de forma robusta los días de la semana de un texto libre (ej: 'Lu, Mi y Vi 15:30').
    Maneja variaciones como Lu, Lun, Lunes, MI, miercoles, V, viernes, sab, etc.
    """
    if not texto or not str(texto).strip() or str(texto).lower() == "nan":
        return ""
    t = quitar_tildes(str(texto).lower())
    # Separar letras y números (ej: viernes10 -> viernes 10, j15 -> j 15)
    t = re.sub(r"([a-z])(\d)", r"\1 \2", t)
    t = re.sub(r"[^a-z0-9\s]", " ", t)

    patrones = [
        ("Lunes", r"\b(lunes|lun|lu|l)\b"),
        ("Martes", r"\b(martes|mar|ma|m)\b"),
        ("Miércoles", r"\b(miercoles|m.ercoles|mier|mie|mi)\b"),
        ("Jueves", r"\b(jueves|jue|ju|j)\b"),
        ("Viernes", r"\b(viernes|vier|vie|vi|v)\b"),
        ("Sábado", r"\b(sabados?|s.bados?|sabdo|sabao|sab|sa)\b"),
        ("Domingo", r"\b(domingos?|dom|do)\b"),
    ]

    encontrados = []
    for dia_nombre, pat in patrones:
        m = re.search(pat, t)
        if m:
            encontrados.append((m.start(), dia_nombre))

    if not encontrados:
        return str(texto).strip()

    encontrados.sort(key=lambda x: x[0])
    dias_ordenados = []
    for _, d in encontrados:
        if d not in dias_ordenados:
            dias_ordenados.append(d)

    return ", ".join(dias_ordenados)


def extraer_datos_alumnos_dataframe(df):
    """
    Extrae la lista de nombres y un diccionario de detalles (jubilado, días de clase)
    a partir del DataFrame de alumnos.
    Columna 1: Nombre (índice 1)
    Columna 2: Jubilado/a (índice 2)
    Columna 5: Días y horarios de clase (índice 5)
    """
    if df.shape[1] < 2:
        return [], {}

    nombres_limpios = []
    detalles = {}
    vistos = set()

    for idx, row in df.iterrows():
        val_nom = row.iloc[1] if df.shape[1] > 1 else None
        if pd.isna(val_nom):
            continue
        nombre_str = str(val_nom).strip()
        if not nombre_str or nombre_str.lower() == "nan":
            continue

        norm = quitar_tildes(nombre_str).lower().strip()
        if norm in ("nombre", "nombres", "alumno", "alumnos", "apellido y nombre", "nombre y apellido", "socio"):
            continue

        # Columna 2 (índice 2): jubilado/a
        es_jubilado = False
        if df.shape[1] > 2:
            val_col2 = row.iloc[2]
            if not pd.isna(val_col2):
                str_col2 = quitar_tildes(str(val_col2).lower().strip())
                if "jubila" in str_col2:
                    es_jubilado = True

        # Columna 5 (índice 5): días de clase
        dias_texto = ""
        horario_raw = ""
        if df.shape[1] > 5:
            val_col5 = row.iloc[5]
            if not pd.isna(val_col5):
                horario_raw = str(val_col5).strip()
                dias_texto = extraer_dias_clase(horario_raw)

        fmt_nom = formatear_nombre_propio(nombre_str)
        if norm not in vistos:
            vistos.add(norm)
            nombres_limpios.append(fmt_nom)

        detalles[norm] = {
            "nombre": fmt_nom,
            "es_jubilado": es_jubilado,
            "dias_clase": dias_texto,
            "horario_raw": horario_raw,
        }

    nombres_limpios.sort(key=lambda x: quitar_tildes(x))
    return nombres_limpios, detalles


def _extraer_segunda_columna_dataframe(df):
    """Extrae y limpia los nombres de la 2da columna (índice 1) de un DataFrame."""
    nombres, _ = extraer_datos_alumnos_dataframe(df)
    return nombres


def cargar_datos_alumnos_local(ruta_alumnos):
    """Lee nombres y detalles desde un archivo Excel local sin modificarlo."""
    if not ruta_alumnos or not os.path.exists(ruta_alumnos):
        return [], {}
    df = pd.read_excel(ruta_alumnos, sheet_name=0, header=None)
    return extraer_datos_alumnos_dataframe(df)


def cargar_nombres_desde_archivo_local(ruta_alumnos):
    """Lee la 2da columna de la 1ra hoja desde un archivo local sin modificarlo."""
    nombres, _ = cargar_datos_alumnos_local(ruta_alumnos)
    return nombres


def _es_binario_excel(datos_bytes):
    """Verifica si los bytes descargados corresponden a un archivo Excel (.xlsx o .xls)."""
    if not datos_bytes or len(datos_bytes) < 8:
        return False
    return datos_bytes[:4] == b"PK\x03\x04" or datos_bytes[:4] == b"\xd0\xcf\x11\xe0"


def _generar_urls_descarga_onedrive(url_original):
    """Genera variantes de URL de descarga directa para enlaces en la nube."""
    url = url_original.strip()
    candidatas = []

    m_gsheet = re.search(r"docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if m_gsheet:
        doc_id = m_gsheet.group(1)
        candidatas.append(f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx")
        return candidatas

    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    candidatas.append(f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content")

    if "download=1" not in url:
        sep = "&" if "?" in url else "?"
        candidatas.append(f"{url}{sep}download=1")

    if "sharepoint.com" in url.lower() and "?" in url:
        base_sin_query = url.split("?")[0]
        candidatas.append(f"{base_sin_query}?download=1")

    if "onedrive.live.com" in url.lower():
        url_dl = re.sub(r"/(view|edit)\.aspx", "/download", url, flags=re.IGNORECASE)
        url_dl = url_dl.replace("onedrive.live.com/redir?", "onedrive.live.com/download?")
        if url_dl not in candidatas:
            candidatas.append(url_dl)

    candidatas.append(url)
    return candidatas


def descargar_datos_alumnos_url(url_compartida):
    """Descarga en memoria el Excel desde OneDrive/SharePoint/Drive y extrae nombres y detalles."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "*/*",
    }

    urls_a_probar = _generar_urls_descarga_onedrive(url_compartida)

    try:
        req_head = urllib.request.Request(url_compartida.strip(), headers=headers)
        with urllib.request.urlopen(req_head, timeout=12) as resp:
            url_final = resp.geturl()
            body_inicial = resp.read()
            if _es_binario_excel(body_inicial):
                df = pd.read_excel(io.BytesIO(body_inicial), sheet_name=0, header=None)
                return extraer_datos_alumnos_dataframe(df)

            if url_final and url_final != url_compartida.strip():
                for extra_u in _generar_urls_descarga_onedrive(url_final):
                    if extra_u not in urls_a_probar:
                        urls_a_probar.insert(0, extra_u)

                parsed = urllib.parse.urlparse(url_final)
                qs = urllib.parse.parse_qs(parsed.query)
                if "resid" in qs:
                    resid = qs["resid"][0]
                    authkey = qs.get("authkey", [""])[0]
                    dl_directo = f"https://onedrive.live.com/download?resid={resid}"
                    if authkey:
                        dl_directo += f"&authkey={authkey}"
                    urls_a_probar.insert(0, dl_directo)

            texto_html = body_inicial.decode("utf-8", errors="ignore")
            m_dl = re.search(r'"@content\.downloadUrl"\s*:\s*"([^"]+)"', texto_html) or re.search(
                r'"downloadUrl"\s*:\s*"([^"]+)"', texto_html
            )
            if m_dl:
                urls_a_probar.insert(0, m_dl.group(1).replace("\\u0026", "&"))
    except Exception:
        pass

    errores = []
    for cand in urls_a_probar:
        try:
            req = urllib.request.Request(cand, headers=headers)
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = resp.read()
                if _es_binario_excel(data):
                    df = pd.read_excel(io.BytesIO(data), sheet_name=0, header=None)
                    nombres, detalles = extraer_datos_alumnos_dataframe(df)
                    if nombres:
                        return nombres, detalles
        except Exception as e:
            errores.append(f"{cand[:40]}... -> {e}")

    raise RuntimeError(
        "No se pudo descargar el archivo Excel automáticamente desde el link de OneDrive.\n\n"
        "Verificá que el enlace tenga permisos de 'Cualquier persona con el link puede ver'."
    )


def descargar_nombres_desde_url(url_compartida):
    """Descarga en memoria el Excel desde OneDrive/SharePoint/Drive y extrae nombres."""
    nombres, _ = descargar_datos_alumnos_url(url_compartida)
    return nombres


def cargar_cache_alumnos():
    """Carga los alumnos guardados previamente en caché local."""
    if os.path.exists(ARCHIVO_CACHE_ALUMNOS):
        try:
            with open(ARCHIVO_CACHE_ALUMNOS, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "detalles" not in data:
                    data["detalles"] = {}
                return data
        except Exception:
            pass
    return {"fecha_sync": "", "hora_sync": "", "nombres": [], "detalles": {}}


def guardar_cache_alumnos(nombres, detalles=None):
    """Guarda la lista de nombres y detalles en cache_alumnos.json con la marca temporal actual."""
    ahora = datetime.now()
    datos = {
        "fecha_sync": ahora.strftime("%Y-%m-%d"),
        "hora_sync": ahora.strftime("%d/%m/%Y %H:%M"),
        "nombres": sorted(list(set(nombres)), key=lambda x: quitar_tildes(x)),
        "detalles": detalles or {},
    }
    try:
        with open(ARCHIVO_CACHE_ALUMNOS, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return datos
