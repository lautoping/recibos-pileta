import os
import io
import re
import json
import base64
import difflib
import urllib.request
import urllib.parse
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook

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


def _grupos_desde_fuente(fuente, nombres):
    """Lee los grupos familiares (2da hoja) desde una ruta o bytes. Si falla, devuelve [] sin cortar la carga."""
    try:
        wb = load_workbook(io.BytesIO(fuente) if isinstance(fuente, (bytes, bytearray)) else fuente)
        return extraer_grupos_familiares(wb, nombres)
    except Exception as e:
        print(f"Aviso: no se pudieron leer los grupos familiares: {e}")
        return []


def cargar_alumnos_completo_local(ruta_alumnos):
    """(nombres, detalles, grupos familiares) desde un archivo local, sin modificarlo."""
    if not ruta_alumnos or not os.path.exists(ruta_alumnos):
        return [], {}, []
    nombres, detalles = cargar_datos_alumnos_local(ruta_alumnos)
    return nombres, detalles, _grupos_desde_fuente(ruta_alumnos, nombres)


def _procesar_bytes_excel(data):
    df = pd.read_excel(io.BytesIO(data), sheet_name=0, header=None)
    nombres, detalles = extraer_datos_alumnos_dataframe(df)
    return nombres, detalles, _grupos_desde_fuente(data, nombres)


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
    nombres, detalles, _ = descargar_alumnos_completo_url(url_compartida)
    return nombres, detalles


def descargar_alumnos_completo_url(url_compartida):
    """Como descargar_datos_alumnos_url, pero también devuelve los grupos familiares de la 2da hoja."""
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
                return _procesar_bytes_excel(body_inicial)

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
                    nombres, detalles, grupos = _procesar_bytes_excel(data)
                    if nombres:
                        return nombres, detalles, grupos
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
                if "grupos_familiares" not in data:
                    data["grupos_familiares"] = []
                return data
        except Exception:
            pass
    return {"fecha_sync": "", "hora_sync": "", "nombres": [], "detalles": {}, "grupos_familiares": []}


def guardar_cache_alumnos(nombres, detalles=None, grupos_familiares=None):
    """Guarda la lista de nombres y detalles en cache_alumnos.json con la marca temporal actual."""
    ahora = datetime.now()
    datos = {
        "fecha_sync": ahora.strftime("%Y-%m-%d"),
        "hora_sync": ahora.strftime("%d/%m/%Y %H:%M"),
        "nombres": sorted(list(set(nombres)), key=lambda x: quitar_tildes(x)),
        "detalles": detalles or {},
        "grupos_familiares": grupos_familiares or [],
    }
    try:
        with open(ARCHIVO_CACHE_ALUMNOS, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return datos


# ==========================================================
# GRUPOS FAMILIARES (2da hoja de la planilla de alumnos)
# ==========================================================
# En la hoja de grupos cada grupo familiar está encerrado con bordes (arriba del primer
# integrante y abajo del último), con los nombres en la 2da columna (B). Hay secciones por
# profesor ("MICA", "AGUS", "ALE") con su propio encabezado. El número de la columna A no
# siempre está en la primera fila del grupo, por eso los grupos se arman con los bordes.

_ENCABEZADOS_GF = ("nombre", "nombres", "apellido y nombre", "nombre y apellido", "integrantes")


def _hoja_grupos(wb):
    for ws in wb.worksheets:
        t = quitar_tildes(ws.title)
        if "grupo" in t or "fliar" in t or "familiar" in t or t.startswith("gf"):
            return ws
    return wb.worksheets[1] if len(wb.worksheets) > 1 else None


def _separar_nombres(texto):
    """'alma, ivan y Franco Fernandez' -> ['Alma Fernandez', 'Ivan Fernandez', 'Franco Fernandez']."""
    limpio = " ".join(str(texto).replace("\n", " ").split())
    partes = [p.strip() for p in re.split(r",|\s+y\s+|\s+e\s+", limpio) if p.strip()]
    if len(partes) <= 1:
        return [formatear_nombre_propio(limpio)] if limpio else []
    ultimo = partes[-1].split()
    apellido = " ".join(ultimo[1:]) if len(ultimo) > 1 else ""
    nombres = []
    for p in partes:
        if apellido and len(p.split()) == 1:
            p = f"{p} {apellido}"
        nombres.append(formatear_nombre_propio(p))
    return nombres


def _tokens(nombre):
    return [t for t in re.split(r"[^a-z0-9ñ]+", quitar_tildes(nombre)) if len(t) > 1]


def buscar_alumno_en_lista(nombre, nombres_alumnos):
    """
    Busca a qué alumno de la lista principal corresponde un nombre escrito a mano en la hoja de GF.
    Coincide si todas las palabras de uno están en el otro (ej. 'Ana Miriam Moyano' ~ 'Ana Moyano').
    Si hay más de un candidato, devuelve None para no adivinar.
    """
    tk_n = set(_tokens(nombre))
    if len(tk_n) < 2:
        return None
    exactos = [a for a in nombres_alumnos if quitar_tildes(a) == quitar_tildes(nombre)]
    if exactos:
        return exactos[0]
    candidatos = []
    for a in nombres_alumnos:
        tk_a = set(_tokens(a))
        if len(tk_a) >= 2 and (tk_n <= tk_a or tk_a <= tk_n):
            candidatos.append(a)
    if len(candidatos) == 1:
        return candidatos[0]
    if candidatos:
        return None

    # Errores de tipeo ('Osvlado' / 'Osvaldo', 'Schneemam' / 'Schneemann'): solo si hay un parecido claro
    objetivo = quitar_tildes(nombre)
    puntajes = sorted(
        ((difflib.SequenceMatcher(None, objetivo, quitar_tildes(a)).ratio(), a) for a in nombres_alumnos),
        reverse=True,
    )
    if puntajes and puntajes[0][0] >= 0.88 and (len(puntajes) == 1 or puntajes[0][0] - puntajes[1][0] >= 0.04):
        return puntajes[0][1]
    return None


def extraer_grupos_familiares(wb, nombres_alumnos=None):
    """
    Lee los grupos familiares de un libro de openpyxl.
    Devuelve lista de grupos: {"seccion", "numero", "miembros": [nombres], "nota"}.
    Cada miembro se reemplaza por el nombre exacto de la lista de alumnos cuando se lo encuentra.
    """
    ws = _hoja_grupos(wb)
    if ws is None:
        return []
    nombres_alumnos = nombres_alumnos or []

    grupos = []
    actual = None
    seccion = ""

    def cerrar():
        nonlocal actual
        if actual and actual["miembros"]:
            grupos.append(actual)
        actual = None

    for r in range(1, ws.max_row + 1):
        val_a = ws.cell(row=r, column=1).value
        celda_b = ws.cell(row=r, column=2)
        val_b = celda_b.value
        txt_b = " ".join(str(val_b).split()) if val_b is not None else ""
        txt_a = str(val_a).strip() if val_a is not None else ""

        # Fila de encabezado de sección (ej. "MICA" / "AGUS  | Nombre | Parentesco")
        if (txt_a and not re.fullmatch(r"\d+(\.0)?", txt_a) and len(txt_a) <= 20 and not txt_b) or quitar_tildes(txt_b) in _ENCABEZADOS_GF:
            cerrar()
            if txt_a and not re.fullmatch(r"\d+(\.0)?", txt_a):
                seccion = txt_a
            continue

        if not txt_b:
            # Fila vacía: termina el grupo (salvo que sea una fila intermedia con borde abajo)
            if actual and celda_b.border.bottom.style:
                cerrar()
            elif not any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, 8)):
                cerrar()
            continue

        if celda_b.border.top.style or actual is None:
            cerrar()
            actual = {"seccion": seccion, "numero": "", "miembros": [], "nota": ""}

        if txt_a and re.fullmatch(r"\d+(\.0)?", txt_a) and not actual["numero"]:
            actual["numero"] = txt_a.split(".")[0]

        for nom in _separar_nombres(txt_b):
            encontrado = buscar_alumno_en_lista(nom, nombres_alumnos)
            final = encontrado or nom
            if final not in actual["miembros"]:
                actual["miembros"].append(final)

        nota = " ".join(
            " ".join(str(ws.cell(row=r, column=c).value).split())
            for c in range(4, 8)
            if ws.cell(row=r, column=c).value is not None and not isinstance(ws.cell(row=r, column=c).value, datetime)
        )
        if nota:
            actual["nota"] = f"{actual['nota']} {nota}".strip()

        if celda_b.border.bottom.style:
            cerrar()

    cerrar()
    return grupos


def indice_grupos(grupos):
    """{nombre normalizado: índice del grupo} para buscar rápido el grupo de un alumno."""
    idx = {}
    for i, g in enumerate(grupos):
        for m in g.get("miembros", []):
            idx.setdefault(quitar_tildes(m), i)
    return idx
