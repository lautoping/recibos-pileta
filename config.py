import os
import json
from datetime import date

CARPETA_BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CONFIG = os.path.join(CARPETA_BASE, "config_recibos.json")
ARCHIVO_CACHE_ALUMNOS = os.path.join(CARPETA_BASE, "cache_alumnos.json")
ARCHIVO_PRECIOS_DEFAULT = os.path.join(CARPETA_BASE, "precios.xlsx")
ARCHIVO_RECIBOS_DEFAULT = os.path.join(CARPETA_BASE, "recibos.xlsx")

COLUMNAS_RECIBOS = [
    "Nro Recibo",
    "Fecha Emisión",
    "Nombre",
    "Actividad",
    "Detalle Actividad",
    "Clases por Semana",
    "Fecha Vencimiento",
    "Inscripción",
    "Descuento",
    "Método de Pago",
    "Nro Carnet",
    "Monto Cuota",
    "Monto Inscripción",
    "Total",
    "Monto Pagado",
    "Saldo Pendiente",
    "Estado Pago",
    "Observaciones",
]


def _es_ruta_temporal(ruta):
    """Verifica si una ruta apunta a un directorio o archivo temporal."""
    if not ruta:
        return False
    ruta_norm = os.path.normpath(str(ruta)).lower()
    base = os.path.basename(ruta_norm)
    if "appdata\\local\\temp" in ruta_norm or "\\temp\\" in ruta_norm or base.startswith("tmp"):
        return True
    return False


def cargar_configuracion():
    """Carga config_recibos.json o devuelve valores por defecto."""
    cfg = {
        "archivo_recibos": ARCHIVO_RECIBOS_DEFAULT,
        "archivo_precios": ARCHIVO_PRECIOS_DEFAULT,
        "url_onedrive_alumnos": "",
        "archivo_alumnos": "",
        "frecuencia_sync_alumnos": "diaria",
        "ultima_fecha_emision": date.today().strftime("%d/%m/%Y"),
    }
    if os.path.exists(ARCHIVO_CONFIG):
        try:
            with open(ARCHIVO_CONFIG, "r", encoding="utf-8") as f:
                guardado = json.load(f)
                cfg.update(guardado)
        except Exception:
            pass

    # Si por alguna prueba o sincronización quedó apuntando a un archivo temporal o inválido,
    # restaurar inmediatamente al archivo real de recibos.xlsx
    if _es_ruta_temporal(cfg.get("archivo_recibos")) or not os.path.exists(cfg.get("archivo_recibos", "")):
        cfg["archivo_recibos"] = ARCHIVO_RECIBOS_DEFAULT
        guardar_configuracion(cfg)

    return cfg


def guardar_configuracion(cfg):
    """Guarda la configuración en config_recibos.json (evitando guardar rutas temporales)."""
    try:
        cfg_a_guardar = dict(cfg)
        if _es_ruta_temporal(cfg_a_guardar.get("archivo_recibos")):
            cfg_a_guardar["archivo_recibos"] = ARCHIVO_RECIBOS_DEFAULT

        with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(cfg_a_guardar, f, indent=2, ensure_ascii=False)
    except Exception:
        pass
