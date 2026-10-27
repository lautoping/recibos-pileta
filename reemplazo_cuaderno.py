import os
import sys
import threading
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, messagebox

# Asegurar que el directorio de este script esté en sys.path para importaciones directas
DIR_ACTUAL = os.path.dirname(os.path.abspath(__file__))
if DIR_ACTUAL not in sys.path:
    sys.path.insert(0, DIR_ACTUAL)

from config import cargar_configuracion, guardar_configuracion
from utils import (
    formatear_nombre_propio,
    parsear_fecha,
    sumar_un_mes,
    formatear_moneda,
)
from excel_manager import (
    cargar_tabla_precios,
    obtener_estado_recibos,
    guardar_fila_recibo,
    eliminar_recibo_por_numero,
)
from alumnos_service import (
    cargar_cache_alumnos,
    guardar_cache_alumnos,
    descargar_datos_alumnos_url,
    cargar_datos_alumnos_local,
    descargar_nombres_desde_url,
    cargar_nombres_desde_archivo_local,
)
from ui_components import (
    AlumnoItem,
    DialogoConfigAlumnos,
    DialogoPagoPendiente,
    DialogoCarnetSinRecibo,
    abrir_calendario,
)


class ReemplazoCuadernoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestión y Carga de Recibos - Pileta")
        self.root.geometry("1020x840")
        self.root.minsize(920, 580)

        self.config = cargar_configuracion()
        self.precios_data = {}
        self.lista_alumnos = []
        self.detalles_alumnos = {}
        self.alumnos_items = []
        self.proximo_recibo_auto = 1
        self.proximo_carnet_auto = 1
        self._sincronizando_pago = False

        self._inicializar_variables()
        self._configurar_estilos()
        self._construir_interfaz()
        self._recargar_todo(inicial=True)

    def _guardar_config(self):
        guardar_configuracion(self.config)

    def _inicializar_variables(self):
        self.var_nro_recibo = tk.StringVar(value="1")
        self.var_editar_nro_recibo = tk.BooleanVar(value=False)

        self.var_fecha_emision = tk.StringVar(
            value=self.config.get("ultima_fecha_emision", date.today().strftime("%d/%m/%Y"))
        )
        self.var_fecha_vencimiento = tk.StringVar()
        self.var_metodo_pago = tk.StringVar(value="Transferencia")  # "Transferencia", "Débito", "Efectivo", "Mixto", "No figura"
        self.var_tipo_mixto = tk.StringVar(value="Efectivo + Transferencia")  # "Efectivo + Transferencia", "Efectivo + Débito", "Transferencia + Débito"
        self.var_monto_efectivo_mixto = tk.StringVar(value="")
        self.var_metodo_resto_mixto = tk.StringVar(value="Transferencia")

        self.var_editar_total_manual = tk.BooleanVar(value=False)
        self.var_total_manual = tk.StringVar(value="0")

        # Variables para Estado del Pago y Pagos Parciales / Deuda:
        self.var_estado_pago = tk.StringVar(value="Pagado")  # "Pagado", "Parcial", "Debe"
        self.var_monto_abonado = tk.StringVar(value="")

        # Variable para Pago Fuera de Término / Vencido (sin descuento):
        self.var_pago_vencido = tk.BooleanVar(value=False)

        self.var_observaciones = tk.StringVar()

        self.var_fecha_emision.trace_add("write", self._al_cambiar_fecha_emision)
        self.var_fecha_vencimiento.trace_add("write", self._al_cambiar_fecha_vencimiento_general)
        self.var_metodo_pago.trace_add("write", lambda *args: self._al_cambiar_metodo_pago())
        self.var_tipo_mixto.trace_add("write", lambda *args: self._al_cambiar_tipo_mixto())
        self.var_monto_efectivo_mixto.trace_add("write", lambda *args: self._actualizar_calculo_precio())
        self.var_metodo_resto_mixto.trace_add("write", lambda *args: self._actualizar_calculo_precio())
        self.var_editar_total_manual.trace_add("write", lambda *args: self._actualizar_calculo_precio())
        self.var_monto_abonado.trace_add("write", lambda *args: self._actualizar_calculo_precio())
        self.var_pago_vencido.trace_add("write", lambda *args: self._actualizar_calculo_precio())

    def _configurar_estilos(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass

        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"), foreground="#0f172a")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 9), foreground="#475569")
        style.configure("Field.TLabel", font=("Segoe UI", 9, "bold"), foreground="#1e293b")
        style.configure("Small.TLabel", font=("Segoe UI", 8), foreground="#64748b")
        style.configure("Save.TButton", font=("Segoe UI", 10, "bold"), padding=8)
        style.configure("BigTotal.TLabel", font=("Segoe UI", 16, "bold"), foreground="#15803d")

    def _al_redimensionar_canvas(self, event):
        ancho = max(event.width, 920)
        self.canvas.itemconfig(self.canvas_window_id, width=ancho)

    def _al_scrollear_canvas(self, *args):
        for it in self.alumnos_items:
            it.autocomplete.cerrar_popup()
        self.canvas.yview(*args)

    def _on_mousewheel(self, event):
        try:
            w = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            w = None

        if w:
            w_class = w.winfo_class()
            if w_class in ("Listbox", "Treeview") or "treeview" in str(w).lower():
                return
            if isinstance(w, (ttk.Scrollbar, tk.Scrollbar)):
                return

        for it in self.alumnos_items:
            it.autocomplete.cerrar_popup()

        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _construir_interfaz(self):
        # Canvas principal con Scrollbar vertical para habilitar scroll fluido en toda la pantalla
        self.canvas = tk.Canvas(self.root, borderwidth=0, highlightthickness=0)
        self.scrollbar_principal = ttk.Scrollbar(self.root, orient="vertical", command=self._al_scrollear_canvas)

        contenedor = ttk.Frame(self.canvas, padding="14 10 14 10")
        self.canvas_window_id = self.canvas.create_window((0, 0), window=contenedor, anchor="nw")

        contenedor.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind("<Configure>", self._al_redimensionar_canvas)
        self.canvas.configure(yscrollcommand=self.scrollbar_principal.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_principal.pack(side=tk.RIGHT, fill=tk.Y)

        self.root.bind_all("<MouseWheel>", self._on_mousewheel)

        # ----------------------------------------------------------
        # 1. ENCABEZADO Y ARCHIVOS VINCULADOS
        # ----------------------------------------------------------
        frame_top = ttk.Frame(contenedor)
        frame_top.pack(fill=tk.X, pady=(0, 4))

        frame_titulos = ttk.Frame(frame_top)
        frame_titulos.pack(side=tk.LEFT)
        ttk.Label(frame_titulos, text="Carga de Recibos - Pileta", style="Title.TLabel").pack(anchor=tk.W)
        self.lbl_subtitulo_precios = ttk.Label(
            frame_titulos,
            text="Numeración correlativa automática, cálculo de precios y base de alumnos.",
            style="Subtitle.TLabel",
        )
        self.lbl_subtitulo_precios.pack(anchor=tk.W)

        frame_acciones_archivos = ttk.Frame(frame_top)
        frame_acciones_archivos.pack(side=tk.RIGHT)

        ttk.Button(
            frame_acciones_archivos,
            text="💵 Pago Pendiente...",
            command=self._abrir_dialogo_pago_pendiente,
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            frame_acciones_archivos,
            text="🪪 Carnet sin Recibo...",
            command=self._abrir_dialogo_carnet_sin_recibo,
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            frame_acciones_archivos,
            text="🔗 Configurar Link Alumnos (OneDrive)...",
            command=self._abrir_dialogo_alumnos,
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            frame_acciones_archivos,
            text="💲 Abrir Precios",
            command=lambda: self._abrir_archivo(self.config["archivo_precios"]),
        ).pack(side=tk.LEFT, padx=3)

        ttk.Button(
            frame_acciones_archivos,
            text="🔄 Sincronizar Ahora",
            command=lambda: self._sincronizar_alumnos_en_segundo_plano(forzar=True),
        ).pack(side=tk.LEFT, padx=3)

        self.lbl_estado_alumnos = ttk.Label(contenedor, text="", style="Small.TLabel")
        self.lbl_estado_alumnos.pack(anchor=tk.W, pady=(0, 6))

        # ----------------------------------------------------------
        # 2. FORMULARIO PRINCIPAL DE CARGA
        # ----------------------------------------------------------
        frame_form = ttk.LabelFrame(contenedor, text=" Nuevo Recibo ", padding="12 8 12 8")
        frame_form.pack(fill=tk.X, pady=(0, 8))

        # --- FILA 1: N° Recibo | Fecha Emisión | Vencimiento | Método de Pago ---
        f_fila1 = ttk.Frame(frame_form)
        f_fila1.pack(fill=tk.X, pady=(0, 8))

        f_recibo = ttk.Frame(f_fila1)
        f_recibo.pack(side=tk.LEFT, padx=(0, 18))
        ttk.Label(f_recibo, text="N° de Recibo:", style="Field.TLabel").pack(anchor=tk.W)
        f_rec_in = ttk.Frame(f_recibo)
        f_rec_in.pack(anchor=tk.W, pady=(2, 0))
        self.entry_nro_recibo = ttk.Entry(
            f_rec_in,
            textvariable=self.var_nro_recibo,
            width=9,
            font=("Segoe UI", 11, "bold"),
            justify="center",
            state="readonly",
        )
        self.entry_nro_recibo.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Checkbutton(
            f_rec_in,
            text="Editar N°",
            variable=self.var_editar_nro_recibo,
            command=self._toggle_editar_nro_recibo,
        ).pack(side=tk.LEFT)

        f_emision = ttk.Frame(f_fila1)
        f_emision.pack(side=tk.LEFT, padx=(0, 18))
        ttk.Label(f_emision, text="Fecha de Emisión:", style="Field.TLabel").pack(anchor=tk.W)
        f_emi_in = ttk.Frame(f_emision)
        f_emi_in.pack(anchor=tk.W, pady=(2, 0))
        self.entry_fecha_emision = ttk.Entry(
            f_emi_in,
            textvariable=self.var_fecha_emision,
            width=11,
            font=("Segoe UI", 10),
            justify="center",
        )
        self.entry_fecha_emision.pack(side=tk.LEFT, padx=(0, 2))
        self.btn_cal_emision = ttk.Button(
            f_emi_in,
            text="📅",
            width=3,
            command=lambda: abrir_calendario(self.btn_cal_emision, self.var_fecha_emision, permitir_vacio=False),
        )
        self.btn_cal_emision.pack(side=tk.LEFT, padx=(0, 4))
        self.entry_fecha_emision.bind(
            "<Double-Button-1>",
            lambda e: abrir_calendario(self.btn_cal_emision, self.var_fecha_emision, permitir_vacio=False),
        )
        ttk.Button(
            f_emi_in,
            text="Hoy",
            width=5,
            command=lambda: self.var_fecha_emision.set(date.today().strftime("%d/%m/%Y")),
        ).pack(side=tk.LEFT)

        f_venc = ttk.Frame(f_fila1)
        f_venc.pack(side=tk.LEFT, padx=(0, 22))
        ttk.Label(f_venc, text="Fecha de Vencimiento (+1 mes):", style="Field.TLabel").pack(anchor=tk.W)
        f_ven_in = ttk.Frame(f_venc)
        f_ven_in.pack(anchor=tk.W, pady=(2, 0))
        self.entry_fecha_vencimiento = ttk.Entry(
            f_ven_in,
            textvariable=self.var_fecha_vencimiento,
            width=11,
            font=("Segoe UI", 10),
            justify="center",
        )
        self.entry_fecha_vencimiento.pack(side=tk.LEFT, padx=(0, 2))
        self.btn_cal_vencimiento = ttk.Button(
            f_ven_in,
            text="📅",
            width=3,
            command=lambda: abrir_calendario(self.btn_cal_vencimiento, self.var_fecha_vencimiento, permitir_vacio=True),
        )
        self.btn_cal_vencimiento.pack(side=tk.LEFT, padx=(0, 4))
        self.entry_fecha_vencimiento.bind(
            "<Double-Button-1>",
            lambda e: abrir_calendario(self.btn_cal_vencimiento, self.var_fecha_vencimiento, permitir_vacio=True),
        )
        ttk.Button(
            f_ven_in,
            text="+1 Mes",
            width=7,
            command=self._forzar_vencimiento_un_mes,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            f_ven_in,
            text="Vacío",
            width=5,
            command=lambda: self.var_fecha_vencimiento.set(""),
        ).pack(side=tk.LEFT)

        # Método de Pago del Recibo
        f_pago = ttk.Frame(f_fila1)
        f_pago.pack(side=tk.LEFT)
        ttk.Label(f_pago, text="Método de Pago del Recibo:", style="Field.TLabel").pack(anchor=tk.W)
        f_pag_opts = ttk.Frame(f_pago)
        f_pag_opts.pack(anchor=tk.W, pady=(2, 0))
        for mp in ("Transferencia", "Débito", "Efectivo", "Mixto", "No figura"):
            ttk.Radiobutton(
                f_pag_opts,
                text=mp,
                value=mp,
                variable=self.var_metodo_pago,
            ).pack(side=tk.LEFT, padx=(0, 8))

        # Condición / Fuera de término (vencido)
        f_cond = ttk.Frame(f_fila1)
        f_cond.pack(side=tk.LEFT, padx=(18, 0))
        ttk.Label(f_cond, text="Condición Vencimiento:", style="Field.TLabel").pack(anchor=tk.W)
        f_cond_in = ttk.Frame(f_cond)
        f_cond_in.pack(anchor=tk.W, pady=(2, 0))
        self.chk_pago_vencido = ttk.Checkbutton(
            f_cond_in,
            text="⏰ Fuera de término (Vencido - Sin dto.)",
            variable=self.var_pago_vencido,
            command=self._al_cambiar_pago_vencido,
        )
        self.chk_pago_vencido.pack(side=tk.LEFT)

        # Subframe para detalle de Pago Mixto (Efectivo / Transferencia / Débito)
        self.f_mixto_detalle = ttk.Frame(frame_form)
        ttk.Label(self.f_mixto_detalle, text="Combinación Mixto:", style="Field.TLabel").pack(side=tk.LEFT, padx=(0, 4))
        for opt_comb in ("Efectivo + Transferencia", "Efectivo + Débito", "Transferencia + Débito"):
            ttk.Radiobutton(
                self.f_mixto_detalle,
                text=opt_comb,
                value=opt_comb,
                variable=self.var_tipo_mixto,
            ).pack(side=tk.LEFT, padx=(0, 6))

        self.lbl_monto_mixto_p1 = ttk.Label(self.f_mixto_detalle, text="Monto Efectivo $:")
        self.lbl_monto_mixto_p1.pack(side=tk.LEFT, padx=(6, 4))
        self.entry_monto_efectivo_mixto = ttk.Entry(
            self.f_mixto_detalle,
            textvariable=self.var_monto_efectivo_mixto,
            width=11,
            font=("Segoe UI", 10, "bold"),
        )
        self.entry_monto_efectivo_mixto.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_mixto_info = ttk.Label(self.f_mixto_detalle, text="", font=("Segoe UI", 9, "bold"), foreground="#4338ca")
        self.lbl_mixto_info.pack(side=tk.LEFT, padx=(6, 0))

        self.sep_form = ttk.Separator(frame_form, orient="horizontal")
        self.sep_form.pack(fill=tk.X, pady=(2, 6))

        # --- CONTENEDOR DINÁMICO DE ALUMNOS ---
        f_alumnos_header = ttk.Frame(frame_form)
        f_alumnos_header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(
            f_alumnos_header,
            text="Personas / Alumnos incluidos en este recibo:",
            style="Field.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Label(
            f_alumnos_header,
            text="(A partir de la 2da persona se aplica automáticamente descuento Grupo Familiar)",
            style="Small.TLabel",
        ).pack(side=tk.LEFT, padx=(8, 0))

        self.frame_alumnos_contenedor = ttk.Frame(frame_form)
        self.frame_alumnos_contenedor.pack(fill=tk.X, pady=(0, 4))

        # Botón para agregar más personas al recibo
        f_btn_mas_alumno = ttk.Frame(frame_form)
        f_btn_mas_alumno.pack(fill=tk.X, pady=(0, 4))
        self.btn_agregar_alumno = ttk.Button(
            f_btn_mas_alumno,
            text="➕ Agregar otra persona a este recibo (Grupo Familiar - GF)",
            command=self._agregar_alumno,
        )
        self.btn_agregar_alumno.pack(side=tk.LEFT)

        # Inicializar el primer alumno (Titular)
        item_titular = AlumnoItem(
            parent_frame=self.frame_alumnos_contenedor,
            index=1,
            app=self,
            es_titular=True,
        )
        self.alumnos_items.append(item_titular)

        # ----------------------------------------------------------
        # 3. RESUMEN DE PRECIO, ESTADO DE PAGO Y CONFIRMACIÓN
        # ----------------------------------------------------------
        frame_resumen = ttk.LabelFrame(contenedor, text=" Cálculo de Importe, Estado de Pago y Confirmación ", padding="12 6 12 6")
        frame_resumen.pack(fill=tk.X, pady=(0, 8))

        f_fila_cobro1 = ttk.Frame(frame_resumen)
        f_fila_cobro1.pack(fill=tk.X, pady=(0, 4))

        # Desglose y Total
        f_total_izq = ttk.Frame(f_fila_cobro1)
        f_total_izq.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.lbl_desglose = ttk.Label(
            f_total_izq,
            text="Cuota: $0   |   Inscripción: $0",
            font=("Segoe UI", 9),
            foreground="#334155",
        )
        self.lbl_desglose.pack(anchor=tk.W)

        f_total_box = ttk.Frame(f_total_izq)
        f_total_box.pack(anchor=tk.W, pady=(2, 0))

        ttk.Label(f_total_box, text="TOTAL DEL RECIBO:", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        self.lbl_total_calculado = ttk.Label(f_total_box, text="$0", style="BigTotal.TLabel")
        self.lbl_total_calculado.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Checkbutton(
            f_total_box,
            text="Modificar total manual:",
            variable=self.var_editar_total_manual,
            command=self._toggle_total_manual,
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.entry_total_manual = ttk.Entry(
            f_total_box,
            textvariable=self.var_total_manual,
            width=10,
            font=("Segoe UI", 10, "bold"),
            state="disabled",
        )
        self.entry_total_manual.pack(side=tk.LEFT)

        # Botones de Guardar y Abrir Excel a la derecha
        f_botones_der = ttk.Frame(f_fila_cobro1)
        f_botones_der.pack(side=tk.RIGHT)

        self.btn_guardar = ttk.Button(
            f_botones_der,
            text="💾  Guardar Recibo y Pasar al Siguiente",
            style="Save.TButton",
            command=self.guardar_recibo,
        )
        self.btn_guardar.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(
            f_botones_der,
            text="📊 Abrir Excel de Recibos",
            command=lambda: self._abrir_archivo(self.config["archivo_recibos"]),
        ).pack(side=tk.LEFT)

        ttk.Separator(frame_resumen, orient="horizontal").pack(fill=tk.X, pady=(4, 6))

        # --- SECCIÓN ESTADO DEL PAGO (PAGADO / PARCIAL / DEBE TODO) ---
        f_fila_pago_estado = ttk.Frame(frame_resumen)
        f_fila_pago_estado.pack(fill=tk.X, pady=(0, 4))

        f_estado_opts = ttk.Frame(f_fila_pago_estado)
        f_estado_opts.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(f_estado_opts, text="Estado del Pago:", style="Field.TLabel").pack(side=tk.LEFT, padx=(0, 8))

        ttk.Radiobutton(
            f_estado_opts,
            text="Pagado completo",
            value="Pagado",
            variable=self.var_estado_pago,
            command=self._al_cambiar_estado_pago,
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Radiobutton(
            f_estado_opts,
            text="Pago parcial (a medias)",
            value="Parcial",
            variable=self.var_estado_pago,
            command=self._al_cambiar_estado_pago,
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Radiobutton(
            f_estado_opts,
            text="No pagó (debe todo)",
            value="Debe",
            variable=self.var_estado_pago,
            command=self._al_cambiar_estado_pago,
        ).pack(side=tk.LEFT)

        # Monto Abonado y Saldo Info
        f_abono_box = ttk.Frame(f_fila_pago_estado)
        f_abono_box.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(f_abono_box, text="Abonó $:").pack(side=tk.LEFT, padx=(0, 4))
        self.entry_monto_abonado = ttk.Entry(
            f_abono_box,
            textvariable=self.var_monto_abonado,
            width=11,
            font=("Segoe UI", 10, "bold"),
            state="disabled",
        )
        self.entry_monto_abonado.pack(side=tk.LEFT, padx=(0, 12))

        self.lbl_saldo_info = ttk.Label(
            f_abono_box,
            text="✅ Pagado completo (Saldo $0)",
            font=("Segoe UI", 9, "bold"),
            foreground="#15803d",
        )
        self.lbl_saldo_info.pack(side=tk.LEFT)

        # Observaciones
        f_obs = ttk.Frame(frame_resumen)
        f_obs.pack(fill=tk.X, pady=(2, 2))
        ttk.Label(f_obs, text="Observaciones (opcional):", style="Field.TLabel").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Entry(f_obs, textvariable=self.var_observaciones, font=("Segoe UI", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )

        # ----------------------------------------------------------
        # 4. TABLA DE ÚLTIMOS RECIBOS CARGADOS
        # ----------------------------------------------------------
        frame_tabla = ttk.LabelFrame(contenedor, text=" Últimos Recibos Registrados ", padding="8 4 8 6")
        frame_tabla.pack(fill=tk.BOTH, expand=True)

        cols_vista = (
            "Nro Recibo",
            "Fecha Emisión",
            "Nombre",
            "Actividad",
            "Clases/Sem",
            "Vencimiento",
            "Inscripción",
            "Descuento",
            "Pago",
            "Carnet",
            "Total",
            "Abonado",
            "Saldo",
            "Estado",
        )
        self.tree = ttk.Treeview(frame_tabla, columns=cols_vista, show="headings", height=6)
        anchos_cols = {
            "Nro Recibo": 70,
            "Fecha Emisión": 85,
            "Nombre": 160,
            "Actividad": 105,
            "Clases/Sem": 65,
            "Vencimiento": 85,
            "Inscripción": 70,
            "Descuento": 95,
            "Pago": 90,
            "Carnet": 55,
            "Total": 80,
            "Abonado": 80,
            "Saldo": 80,
            "Estado": 75,
        }
        for col in cols_vista:
            self.tree.heading(col, text=col)
            align = "w" if col in ("Nombre", "Actividad") else "center"
            if col in ("Total", "Abonado", "Saldo"):
                align = "e"
            self.tree.column(col, width=anchos_cols[col], anchor=align)

        # Configurar colores de tags para destacar en la tabla de la app
        self.tree.tag_configure("tag_parcial", background="#FEF08A", foreground="#854D0E")
        self.tree.tag_configure("tag_debe", background="#FECACA", foreground="#991B1B")

        scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        f_pie_tabla = ttk.Frame(contenedor)
        f_pie_tabla.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(
            f_pie_tabla,
            text="💵 Cobrar saldo del recibo seleccionado...",
            command=self._abrir_dialogo_pago_seleccionado,
        ).pack(side=tk.LEFT)
        ttk.Button(
            f_pie_tabla,
            text="🗑️ Eliminar recibo seleccionado de la tabla",
            command=self._eliminar_recibo_seleccionado,
        ).pack(side=tk.RIGHT)

    # ==========================================================
    # MANEJO DINÁMICO DE MÚLTIPLES PERSONAS / ALUMNOS
    # ==========================================================
    def _agregar_alumno(self):
        nuevo_idx = len(self.alumnos_items) + 1
        nuevo_item = AlumnoItem(
            parent_frame=self.frame_alumnos_contenedor,
            index=nuevo_idx,
            app=self,
            es_titular=False,
        )
        if self.alumnos_items:
            titular = self.alumnos_items[0]
            act_tit = titular.var_actividad.get()
            if act_tit not in ("Clase Adaptación (CdA)", "Solo Inscripción"):
                nuevo_item.var_actividad.set(act_tit)
                nuevo_item.var_clases_sem.set(titular.var_clases_sem.get())
            else:
                nuevo_item.var_actividad.set("Natación")
                nuevo_item.var_clases_sem.set(2)
            nuevo_item._al_cambiar_actividad()

        self.alumnos_items.append(nuevo_item)
        self._actualizar_calculo_precio()

        # Actualizar región de scroll y enfocar en el nuevo alumno
        self.root.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        nuevo_item.autocomplete.entry.focus_set()
        try:
            bbox = self.canvas.bbox("all")
            y_frame = nuevo_item.frame.winfo_y()
            if bbox and bbox[3] > 0:
                frac = max(0.0, min(1.0, (y_frame - 20) / bbox[3]))
                self.canvas.yview_moveto(frac)
        except Exception:
            pass

    def _quitar_alumno(self, item):
        if item in self.alumnos_items and not item.es_titular:
            item.frame.destroy()
            self.alumnos_items.remove(item)
            for idx, it in enumerate(self.alumnos_items, start=1):
                it.actualizar_indice(idx)
            self._actualizar_calculo_precio()
            self.root.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sugerir_proximo_carnet(self, alumno_solicitante):
        carnets_actuales = []
        for it in self.alumnos_items:
            if it != alumno_solicitante and it.var_usar_carnet.get():
                try:
                    carnets_actuales.append(int(it.var_nro_carnet.get().strip()))
                except ValueError:
                    pass
        if carnets_actuales:
            return max(max(carnets_actuales) + 1, self.proximo_carnet_auto)
        return self.proximo_carnet_auto

    # ==========================================================
    # SINCRONIZACIÓN DE ALUMNOS (LINK ONEDRIVE + CACHÉ DIARIA)
    # ==========================================================
    def _inicializar_alumnos(self):
        cache = cargar_cache_alumnos()
        nombres_cache = cache.get("nombres", [])
        self.detalles_alumnos = cache.get("detalles", {})
        fecha_sync = cache.get("fecha_sync", "")
        hora_sync = cache.get("hora_sync", "")

        if nombres_cache:
            self.lista_alumnos = nombres_cache
            for it in self.alumnos_items:
                it.autocomplete.actualizar_base(self.lista_alumnos)
                it._al_cambiar_nombre()

        url_od = self.config.get("url_onedrive_alumnos", "").strip()
        ruta_local = self.config.get("archivo_alumnos", "").strip()
        frecuencia = self.config.get("frecuencia_sync_alumnos", "diaria")
        hoy_iso = date.today().strftime("%Y-%m-%d")

        if url_od:
            debe_sincronizar = (frecuencia == "siempre") or (fecha_sync != hoy_iso) or (not nombres_cache)
            if debe_sincronizar:
                self._sincronizar_alumnos_en_segundo_plano(forzar=False)
            else:
                self.lbl_estado_alumnos.config(
                    text=f"✅ Base de alumnos al día ({len(self.lista_alumnos)} nombres de 2da columna - última act. hoy {hora_sync})."
                )
        elif ruta_local and os.path.exists(ruta_local):
            try:
                nombres, detalles = cargar_datos_alumnos_local(ruta_local)
                datos_c = guardar_cache_alumnos(nombres, detalles)
                self.lista_alumnos = nombres
                self.detalles_alumnos = detalles
                for it in self.alumnos_items:
                    it.autocomplete.actualizar_base(self.lista_alumnos)
                    it._al_cambiar_nombre()
                self.lbl_estado_alumnos.config(
                    text=f"✅ Base de alumnos local ({len(self.lista_alumnos)} nombres - {datos_c['hora_sync']}): {ruta_local}"
                )
            except Exception as e:
                self.lbl_estado_alumnos.config(text=f"⚠️ Error leyendo archivo local de alumnos: {e}")
        else:
            if nombres_cache:
                self.lbl_estado_alumnos.config(
                    text=f"ℹ️ Usando {len(nombres_cache)} alumnos en caché ({hora_sync}). Configurá el link en '🔗 Configurar Link Alumnos (OneDrive)...'"
                )
            else:
                self.lbl_estado_alumnos.config(
                    text="ℹ️ Hacé clic en '🔗 Configurar Link Alumnos (OneDrive)...' arriba a la derecha para pegar el link compartido."
                )

    def _sincronizar_alumnos_en_segundo_plano(self, forzar=False):
        url_od = self.config.get("url_onedrive_alumnos", "").strip()
        ruta_local = self.config.get("archivo_alumnos", "").strip()

        if not url_od and not ruta_local:
            if forzar:
                self._abrir_dialogo_alumnos()
            return

        self.lbl_estado_alumnos.config(text="⏳ Sincronizando base de alumnos en segundo plano...")

        def tarea():
            try:
                if url_od:
                    nombres, detalles = descargar_datos_alumnos_url(url_od)
                    origen = "OneDrive / Nube"
                else:
                    nombres, detalles = cargar_datos_alumnos_local(ruta_local)
                    origen = "Archivo local"

                datos_guardados = guardar_cache_alumnos(nombres, detalles)
                self.root.after(
                    0,
                    lambda: self._finalizar_sync_alumnos_exito(
                        nombres, detalles, datos_guardados["hora_sync"], origen, forzar
                    ),
                )
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: self._finalizar_sync_alumnos_error(err_msg, forzar))

        threading.Thread(target=tarea, daemon=True).start()

    def _finalizar_sync_alumnos_exito(self, nombres, detalles, hora_str, origen_str, mostrar_popup):
        self.lista_alumnos = nombres
        self.detalles_alumnos = detalles
        for it in self.alumnos_items:
            it.autocomplete.actualizar_base(self.lista_alumnos)
            it._al_cambiar_nombre()
        self.lbl_estado_alumnos.config(
            text=f"✅ {len(nombres)} alumnos importados desde {origen_str} (2da columna, 1ra hoja) - Actualizado: {hora_str}"
        )
        if mostrar_popup:
            messagebox.showinfo(
                "Sincronización Exitosa",
                f"Se importaron {len(nombres)} nombres desde la 2da columna de la primera hoja.\n\n"
                f"Origen: {origen_str}\n"
                f"Fecha: {hora_str}",
            )

    def _finalizar_sync_alumnos_error(self, mensaje_error, mostrar_popup):
        cache = cargar_cache_alumnos()
        n_cache = len(cache.get("nombres", []))
        if n_cache > 0:
            self.lbl_estado_alumnos.config(
                text=f"⚠️ No se pudo conectar al link hoy; usando {n_cache} alumnos guardados en caché ({cache.get('hora_sync', '')})."
            )
        else:
            self.lbl_estado_alumnos.config(
                text="⚠️ No se pudo descargar desde el link de OneDrive. Revisá la configuración del enlace."
            )
        if mostrar_popup:
            messagebox.showerror("Error al Sincronizar Link", mensaje_error)

    def _abrir_dialogo_alumnos(self):
        def al_guardar():
            self._guardar_config()
            self._sincronizar_alumnos_en_segundo_plano(forzar=True)

        DialogoConfigAlumnos(self.root, self.config, on_guardar_callback=al_guardar)

    # ==========================================================
    # CÁLCULOS DE PRECIOS, ESTADO DE PAGO Y EVENTOS
    # ==========================================================
    def _recargar_todo(self, inicial=False):
        self.precios_data = cargar_tabla_precios(self.config["archivo_precios"])
        insc_monto = formatear_moneda(self.precios_data.get("inscripcion", 15000))
        vig = self.precios_data.get("vigencia", "")
        if vig:
            self.lbl_subtitulo_precios.config(
                text=f"Lista de precios activa ({vig})  |  Inscripción: {insc_monto}"
            )

        if inicial:
            self._inicializar_alumnos()

        prox_rec, prox_car, ult_fecha, ultimos = obtener_estado_recibos(self.config["archivo_recibos"])
        self.proximo_recibo_auto = prox_rec
        self.proximo_carnet_auto = prox_car

        if not self.var_editar_nro_recibo.get():
            self.var_nro_recibo.set(str(self.proximo_recibo_auto))

        if inicial:
            fecha_arranque = ult_fecha or self.config.get("ultima_fecha_emision") or date.today().strftime("%d/%m/%Y")
            self.var_fecha_emision.set(fecha_arranque)
            self._forzar_vencimiento_un_mes()

        self._refrescar_tabla(ultimos)
        self._actualizar_calculo_precio()

    def _al_cambiar_fecha_emision(self, *args):
        dt = parsear_fecha(self.var_fecha_emision.get())
        if dt:
            venc = sumar_un_mes(dt)
            venc_str = venc.strftime("%d/%m/%Y")
            self.var_fecha_vencimiento.set(venc_str)
            self._propagar_vencimiento_a_alumnos(venc_str)

    def _forzar_vencimiento_un_mes(self):
        dt = parsear_fecha(self.var_fecha_emision.get())
        if dt:
            venc = sumar_un_mes(dt)
            venc_str = venc.strftime("%d/%m/%Y")
        else:
            venc_str = sumar_un_mes(date.today()).strftime("%d/%m/%Y")
        self.var_fecha_vencimiento.set(venc_str)
        self._propagar_vencimiento_a_alumnos(venc_str)

    def _al_cambiar_fecha_vencimiento_general(self, *args):
        if getattr(self, "_propagando_vencimiento", False):
            return
        self._propagando_vencimiento = True
        try:
            venc_str = self.var_fecha_vencimiento.get()
            self._propagar_vencimiento_a_alumnos(venc_str)
        finally:
            self._propagando_vencimiento = False

    def _propagar_vencimiento_a_alumnos(self, venc_str):
        for item in self.alumnos_items:
            if item.var_actividad.get() != "Clase Adaptación (CdA)":
                item.var_fecha_vencimiento.set(venc_str)

    def _al_cambiar_tipo_mixto(self):
        comb = self.var_tipo_mixto.get()
        if hasattr(self, "lbl_monto_mixto_p1"):
            if comb == "Transferencia + Débito":
                self.lbl_monto_mixto_p1.config(text="Monto Transferencia $:")
            else:
                self.lbl_monto_mixto_p1.config(text="Monto Efectivo $:")
        self._actualizar_calculo_precio()

    def _obtener_efectivo_mixto(self):
        if self.var_metodo_pago.get() != "Mixto":
            return 0.0
        if "Efectivo" not in self.var_tipo_mixto.get():
            return 0.0
        str_ef = self.var_monto_efectivo_mixto.get().strip()
        if not str_ef:
            return 0.0
        try:
            limpio = str_ef.replace("$", "").replace(".", "").replace(",", ".").strip()
            return max(0.0, float(limpio))
        except ValueError:
            return 0.0

    def _obtener_monto_p1_mixto(self):
        str_val = self.var_monto_efectivo_mixto.get().strip()
        if not str_val:
            return 0.0
        try:
            limpio = str_val.replace("$", "").replace(".", "").replace(",", ".").strip()
            return max(0.0, float(limpio))
        except ValueError:
            return 0.0

    def _toggle_editar_nro_recibo(self):
        if self.var_editar_nro_recibo.get():
            self.entry_nro_recibo.config(state="normal")
            self.entry_nro_recibo.focus_set()
        else:
            self.entry_nro_recibo.config(state="readonly")
            self.var_nro_recibo.set(str(self.proximo_recibo_auto))

    def _al_cambiar_metodo_pago(self):
        if self._sincronizando_pago:
            return
        self._sincronizando_pago = True
        try:
            mp = self.var_metodo_pago.get()
            if hasattr(self, "f_mixto_detalle") and hasattr(self, "sep_form"):
                if mp == "Mixto":
                    self.f_mixto_detalle.pack(fill=tk.X, pady=(0, 6), before=self.sep_form)
                else:
                    self.f_mixto_detalle.pack_forget()

            if self.alumnos_items:
                titular = self.alumnos_items[0]
                desc = titular.var_descuento.get()
                if not self.var_pago_vencido.get():
                    if desc not in ("Jubilado", "Grupo familiar"):
                        if mp == "Efectivo" and desc == "Ninguno":
                            titular.var_descuento.set("Efectivo")
                        elif mp in ("Transferencia", "Débito", "No figura") and desc == "Efectivo":
                            titular.var_descuento.set("Ninguno")
        finally:
            self._sincronizando_pago = False
            self._actualizar_calculo_precio()

    def _al_cambiar_pago_vencido(self):
        self._actualizar_calculo_precio()

    def _al_cambiar_estado_pago(self):
        st = self.var_estado_pago.get()
        if st == "Pagado":
            self.entry_monto_abonado.config(state="disabled")
        elif st == "Parcial":
            self.entry_monto_abonado.config(state="normal")
            if not self.var_monto_abonado.get().strip():
                self.entry_monto_abonado.focus_set()
        elif st == "Debe":
            self.entry_monto_abonado.config(state="disabled")
            self.var_monto_abonado.set("0")
        self._actualizar_calculo_precio()

    def _toggle_total_manual(self):
        if self.var_editar_total_manual.get():
            self.entry_total_manual.config(state="normal")
            self.entry_total_manual.focus_set()
        else:
            self.entry_total_manual.config(state="disabled")
            self._actualizar_calculo_precio()

    def _actualizar_calculo_precio(self, *args):
        total_calculado = 0.0
        desglose_parts = []

        for item in self.alumnos_items:
            cuota, insc, subtotal = item.obtener_montos()
            item.actualizar_etiqueta_subtotal()
            total_calculado += subtotal
            nom = item.var_nombre.get().strip() or f"Alumno {item.index}"
            desglose_parts.append(f"{nom}: {formatear_moneda(subtotal)}")

        if self.var_editar_total_manual.get():
            try:
                limpio = self.var_total_manual.get().replace("$", "").replace(".", "").replace(",", ".").strip()
                total_final = float(limpio)
            except ValueError:
                total_final = total_calculado
        else:
            total_final = total_calculado
            self.var_total_manual.set(str(int(total_calculado) if total_calculado.is_integer() else total_calculado))

        if len(self.alumnos_items) > 1:
            self.lbl_desglose.config(text="  |  ".join(desglose_parts))
        elif self.alumnos_items:
            c, i, s = self.alumnos_items[0].obtener_montos()
            desc = self.alumnos_items[0].var_descuento.get()
            clases = self.alumnos_items[0].var_clases_sem.get() * 4
            self.lbl_desglose.config(
                text=f"Cuota ({clases} clases/mes - {desc}): {formatear_moneda(c)}   |   Inscripción: {formatear_moneda(i)}"
            )

        self.lbl_total_calculado.config(text=formatear_moneda(total_final))

        # Si es método de pago Mixto, actualizar la etiqueta de desglose del mixto
        if self.var_metodo_pago.get() == "Mixto" and hasattr(self, "lbl_mixto_info"):
            tipo_m = self.var_tipo_mixto.get()
            monto_p1 = self._obtener_monto_p1_mixto()
            resto = max(0.0, total_final - monto_p1)

            if tipo_m == "Efectivo + Transferencia":
                info_tit = ""
                if self.alumnos_items and not self.var_pago_vencido.get():
                    titular = self.alumnos_items[0]
                    if titular.var_descuento.get() == "Jubilado":
                        info_tit = " | ✅ Dto. Jubilado aplicado"
                    elif titular.var_actividad.get() not in ("Clase Adaptación (CdA)", "Solo Inscripción"):
                        cuota_ef_t = float(self.precios_data.get("cuotas", {}).get(titular.var_clases_sem.get(), {}).get("Efectivo", 0.0))
                        if cuota_ef_t > 0 and monto_p1 >= cuota_ef_t:
                            info_tit = " | ✅ Aplica dto. efectivo al Titular"
                        elif cuota_ef_t > 0:
                            info_tit = f" | ⚠️ Titular requiere mín. {formatear_moneda(cuota_ef_t)} ef."
                self.lbl_mixto_info.config(
                    text=f"Efectivo: {formatear_moneda(monto_p1)}  +  Transferencia: {formatear_moneda(resto)}{info_tit}"
                )
            elif tipo_m == "Efectivo + Débito":
                info_tit = ""
                if self.alumnos_items and not self.var_pago_vencido.get():
                    titular = self.alumnos_items[0]
                    if titular.var_descuento.get() == "Jubilado":
                        info_tit = " | ✅ Dto. Jubilado aplicado"
                    elif titular.var_actividad.get() not in ("Clase Adaptación (CdA)", "Solo Inscripción"):
                        cuota_ef_t = float(self.precios_data.get("cuotas", {}).get(titular.var_clases_sem.get(), {}).get("Efectivo", 0.0))
                        if cuota_ef_t > 0 and monto_p1 >= cuota_ef_t:
                            info_tit = " | ✅ Aplica dto. efectivo al Titular"
                        elif cuota_ef_t > 0:
                            info_tit = f" | ⚠️ Titular requiere mín. {formatear_moneda(cuota_ef_t)} ef."
                self.lbl_mixto_info.config(
                    text=f"Efectivo: {formatear_moneda(monto_p1)}  +  Débito: {formatear_moneda(resto)}{info_tit}"
                )
            else:  # Transferencia + Débito
                info_tit = ""
                if self.alumnos_items and not self.var_pago_vencido.get():
                    titular = self.alumnos_items[0]
                    if titular.var_descuento.get() == "Jubilado":
                        info_tit = " | ✅ Dto. Jubilado aplicado"
                    else:
                        info_tit = " | (Sin dto. efectivo)"
                self.lbl_mixto_info.config(
                    text=f"Transferencia: {formatear_moneda(monto_p1)}  +  Débito: {formatear_moneda(resto)}{info_tit}"
                )

        # Actualizar visualización del estado del pago
        st = self.var_estado_pago.get()
        if st == "Pagado":
            self.lbl_saldo_info.config(
                text=f"✅ Pagado completo: {formatear_moneda(total_final)} (Saldo pendiente: $0)",
                foreground="#15803d",
            )
        elif st == "Parcial":
            try:
                limpio_ab = self.var_monto_abonado.get().replace("$", "").replace(".", "").replace(",", ".").strip()
                val_abonado = float(limpio_ab) if limpio_ab else 0.0
            except ValueError:
                val_abonado = 0.0
            val_abonado = min(max(0.0, val_abonado), total_final)
            saldo = max(0.0, total_final - val_abonado)
            self.lbl_saldo_info.config(
                text=f"⚠️ PAGO PARCIAL: Abonó {formatear_moneda(val_abonado)} — Resta pagar: {formatear_moneda(saldo)} (se marcará en AMARILLO en Excel)",
                foreground="#854d0e",
            )
        elif st == "Debe":
            self.lbl_saldo_info.config(
                text=f"❌ NO PAGÓ (DEBE TODO): {formatear_moneda(total_final)} (se marcará en ROJO en Excel)",
                foreground="#b91c1c",
            )

    def _abrir_archivo(self, ruta):
        if os.path.exists(ruta):
            try:
                os.startfile(ruta)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showwarning("Aviso", f"El archivo aún no existe:\n{ruta}")

    # ==========================================================
    # GUARDADO DE RECIBO Y LIMPIEZA PARA EL SIGUIENTE
    # ==========================================================
    def guardar_recibo(self):
        if not self.alumnos_items:
            messagebox.showwarning("Faltan Datos", "No hay ningún alumno configurado en el recibo.")
            return

        for item in self.alumnos_items:
            nombre = formatear_nombre_propio(item.var_nombre.get())
            if not nombre:
                messagebox.showwarning("Falta Nombre", f"Por favor ingresá el nombre del Alumno {item.index}.")
                item.autocomplete.entry.focus_set()
                return

            actividad = item.var_actividad.get()
            detalle_act = item.var_actividad_otro.get().strip() if actividad == "Otro" else ""
            if actividad == "Otro" and not detalle_act:
                messagebox.showwarning(
                    "Detalle de Actividad",
                    f"El Alumno {item.index} tiene seleccionada la actividad 'Otro'. Especificá qué actividad realiza.",
                )
                item.entry_actividad_otro.focus_set()
                return

            if item.var_usar_carnet.get():
                nro_car_str = item.var_nro_carnet.get().strip()
                if not nro_car_str:
                    messagebox.showwarning(
                        "N° de Carnet",
                        f"Tildaste asignar carnet para el Alumno {item.index} pero el número está vacío.",
                    )
                    item.entry_carnet.focus_set()
                    return

        try:
            nro_recibo = int(self.var_nro_recibo.get().strip())
        except ValueError:
            messagebox.showwarning("N° de Recibo inválido", "El número de recibo debe ser un número entero.")
            return

        dt_emision = parsear_fecha(self.var_fecha_emision.get())
        if not dt_emision:
            messagebox.showwarning("Fecha de Emisión", "Ingresá una Fecha de Emisión válida con formato DD/MM/AAAA.")
            self.entry_fecha_emision.focus_set()
            return

        # Validar vencimiento general (opcional: puede dejarse vacío si falta en el cuaderno o no aplica)
        str_venc_gen = self.var_fecha_vencimiento.get().strip()
        if str_venc_gen:
            dt_venc = parsear_fecha(str_venc_gen)
            if not dt_venc:
                messagebox.showwarning(
                    "Fecha de Vencimiento",
                    "La Fecha de Vencimiento general no es válida. Usá formato DD/MM/AAAA o dejala vacía.",
                )
                self.entry_fecha_vencimiento.focus_set()
                return

        # Validar vencimientos individuales de cada alumno (CdA siempre sin vencimiento)
        fechas_venc_alumnos = []
        for item in self.alumnos_items:
            act_i = item.var_actividad.get()
            if act_i == "Clase Adaptación (CdA)":
                fechas_venc_alumnos.append("")
            else:
                v_str = item.var_fecha_vencimiento.get().strip()
                if v_str:
                    dt_v_item = parsear_fecha(v_str)
                    if not dt_v_item:
                        messagebox.showwarning(
                            "Fecha de Vencimiento Inválida",
                            f"La fecha de vencimiento del Alumno {item.index} ('{v_str}') no es válida.\n"
                            "Debe tener formato DD/MM/AAAA o dejarse vacía.",
                        )
                        item.entry_venc.focus_set()
                        return
                    fechas_venc_alumnos.append(dt_v_item.strftime("%d/%m/%Y"))
                else:
                    fechas_venc_alumnos.append("")

        # Calcular montos y saldo
        estado_pago = self.var_estado_pago.get()  # "Pagado", "Parcial", "Debe"
        monto_cuota_tot = 0.0
        monto_insc_tot = 0.0
        subtotales_alumnos = []

        for item in self.alumnos_items:
            c, i, s = item.obtener_montos()
            monto_cuota_tot += c
            monto_insc_tot += i
            subtotales_alumnos.append((c, i, s))

        total_calculado = monto_cuota_tot + monto_insc_tot

        if self.var_editar_total_manual.get():
            try:
                limpio = self.var_total_manual.get().replace("$", "").replace(".", "").replace(",", ".").strip()
                total_final = float(limpio)
            except ValueError:
                total_final = total_calculado
        else:
            total_final = total_calculado

        total_abonado = 0.0
        if estado_pago == "Pagado":
            total_abonado = total_final
        elif estado_pago == "Debe":
            total_abonado = 0.0
        elif estado_pago == "Parcial":
            str_ab = self.var_monto_abonado.get().strip()
            if not str_ab:
                messagebox.showwarning(
                    "Monto Abonado",
                    "Seleccionaste 'Pago parcial'. Por favor ingresá cuánto dinero abonó el alumno/familia.",
                )
                self.entry_monto_abonado.focus_set()
                return
            try:
                limpio_ab = str_ab.replace("$", "").replace(".", "").replace(",", ".").strip()
                total_abonado = float(limpio_ab)
            except ValueError:
                messagebox.showwarning("Monto inválido", "El monto abonado ingresado no es válido.")
                self.entry_monto_abonado.focus_set()
                return

            if total_abonado < 0:
                messagebox.showwarning("Monto inválido", "El monto abonado no puede ser negativo.")
                return
            if total_abonado >= total_final:
                estado_pago = "Pagado"
                total_abonado = total_final

        saldo_total_recibo = max(0.0, total_final - total_abonado)

        fecha_emision_str = dt_emision.strftime("%d/%m/%Y")
        mp = self.var_metodo_pago.get()
        if mp == "Mixto":
            tipo_m = self.var_tipo_mixto.get()
            monto_p1 = self._obtener_monto_p1_mixto()
            resto_monto = max(0.0, total_final - monto_p1)
            if tipo_m == "Efectivo + Transferencia":
                metodo_pago_str = f"Mixto ({formatear_moneda(monto_p1)} Efectivo / {formatear_moneda(resto_monto)} Transferencia)"
            elif tipo_m == "Efectivo + Débito":
                metodo_pago_str = f"Mixto ({formatear_moneda(monto_p1)} Efectivo / {formatear_moneda(resto_monto)} Débito)"
            else:  # Transferencia + Débito
                metodo_pago_str = f"Mixto ({formatear_moneda(monto_p1)} Transferencia / {formatear_moneda(resto_monto)} Débito)"
        else:
            metodo_pago_str = mp

        user_obs = self.var_observaciones.get().strip()
        n_alumnos = len(self.alumnos_items)

        obs_extra = ""
        if n_alumnos > 1:
            if estado_pago == "Parcial":
                obs_extra = f"Recibo grupal ({n_alumnos} pers) - Abonó {formatear_moneda(total_abonado)} de {formatear_moneda(total_final)} - Saldo {formatear_moneda(saldo_total_recibo)}"
            elif estado_pago == "Debe":
                obs_extra = f"Recibo grupal ({n_alumnos} pers) - Debe todo: {formatear_moneda(total_final)}"
            else:
                obs_extra = f"Recibo grupal ({n_alumnos} pers)"

        obs_completa = f"{user_obs} | {obs_extra}".strip(" |") if obs_extra else user_obs

        filas_a_guardar = []
        restante_abonado = total_abonado
        fuera_termino = self.var_pago_vencido.get()

        for idx, item in enumerate(self.alumnos_items):
            nombre_al = formatear_nombre_propio(item.var_nombre.get())
            actividad_al = item.var_actividad.get()
            detalle_act_al = item.var_actividad_otro.get().strip() if actividad_al == "Otro" else ""

            c_i, ins_i, sub_i = subtotales_alumnos[idx]

            clases_val = "" if actividad_al in ("Clase Adaptación (CdA)", "Solo Inscripción") else item.var_clases_sem.get()
            tipo_insc_al = item.var_inscripcion.get()

            if actividad_al in ("Clase Adaptación (CdA)", "Solo Inscripción"):
                descuento_al = "Ninguno"
            elif fuera_termino:
                descuento_al = "Ninguno (Vencido)"
            else:
                descuento_al = item.var_descuento.get()

            carnet_val = ""
            if item.var_usar_carnet.get():
                c_str = item.var_nro_carnet.get().strip()
                try:
                    carnet_val = int(c_str)
                except ValueError:
                    carnet_val = c_str

            if n_alumnos == 1:
                total_col = total_final
                monto_pagado_i = total_abonado
                saldo_i = saldo_total_recibo
                estado_i = estado_pago
            else:
                # Para grupo familiar: en las columnas Total, Monto Pagado y Saldo Pendiente
                # se deja vacío en los primeros miembros y se asienta el total/abono final en el último miembro
                if idx < n_alumnos - 1:
                    total_col = ""
                    monto_pagado_i = ""
                    saldo_i = ""
                    estado_i = estado_pago
                else:
                    total_col = total_final
                    monto_pagado_i = total_abonado
                    saldo_i = saldo_total_recibo
                    estado_i = estado_pago

            obs_alumno = obs_completa
            if tipo_insc_al in ("Ticket Sochi", "Ticket Malva", "Sochi", "Malva"):
                obs_alumno = f"{obs_alumno} [{tipo_insc_al} presentado - Inscripción $0]".strip()
            elif tipo_insc_al == "CdA":
                obs_alumno = f"{obs_alumno} [CdA previa abonada - Inscripción $0]".strip()

            if fuera_termino and actividad_al not in ("Clase Adaptación (CdA)", "Solo Inscripción"):
                obs_alumno = f"{obs_alumno} [Pago fuera de término - Sin dto.]".strip()

            # Si es una persona sola con descuento de Grupo Familiar, asentar en obs el familiar asociado
            if n_alumnos == 1 and item.var_descuento.get() == "Grupo familiar":
                gf_info = item.var_gf_asociado.get().strip()
                if gf_info:
                    gf_txt = gf_info if gf_info.upper().startswith("GF") else f"GF {gf_info}"
                    obs_alumno = f"{obs_alumno} {gf_txt}".strip() if obs_alumno else gf_txt

            fecha_venc_al = fechas_venc_alumnos[idx]

            fila_datos = {
                "Nro Recibo": nro_recibo,
                "Fecha Emisión": fecha_emision_str,
                "Nombre": nombre_al,
                "Actividad": actividad_al,
                "Detalle Actividad": detalle_act_al,
                "Clases por Semana": clases_val,
                "Fecha Vencimiento": fecha_venc_al,
                "Inscripción": tipo_insc_al,
                "Descuento": descuento_al,
                "Método de Pago": metodo_pago_str,
                "Nro Carnet": carnet_val,
                "Monto Cuota": c_i,
                "Monto Inscripción": ins_i,
                "Total": total_col,
                "Monto Pagado": monto_pagado_i,
                "Saldo Pendiente": saldo_i,
                "Estado Pago": estado_i,
                "Observaciones": obs_alumno,
            }
            filas_a_guardar.append(fila_datos)

        try:
            for fila in filas_a_guardar:
                guardar_fila_recibo(self.config["archivo_recibos"], fila)
        except PermissionError:
            messagebox.showerror(
                "Archivo Excel Abierto",
                "No se pudo guardar el recibo porque 'recibos.xlsx' está abierto en Excel.\n\n"
                "Por favor cerralo en Excel y volvé a presionar Guardar.",
            )
            return
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"Ocurrió un problema guardando el recibo:\n{e}")
            return

        self.config["ultima_fecha_emision"] = fecha_emision_str
        self._guardar_config()

        # Resetear formulario para el siguiente recibo
        while len(self.alumnos_items) > 1:
            it = self.alumnos_items.pop()
            it.frame.destroy()

        titular = self.alumnos_items[0]
        titular.var_nombre.set("")
        titular.var_gf_asociado.set("")
        if hasattr(titular, "lbl_dias_clase") and titular.lbl_dias_clase:
            titular.lbl_dias_clase.config(text="")
        if hasattr(titular, "f_gf_asociado") and titular.f_gf_asociado:
            titular.f_gf_asociado.pack_forget()
        titular.var_actividad.set("Natación")
        titular._al_cambiar_actividad()
        titular.var_clases_sem.set(2)
        titular.var_inscripcion.set("No")
        titular.var_usar_carnet.set(False)
        titular._toggle_carnet()
        mp = self.var_metodo_pago.get()
        titular.var_descuento.set("Efectivo" if mp == "Efectivo" else "Ninguno")
        titular.var_fecha_vencimiento.set(self.var_fecha_vencimiento.get())

        self.var_monto_efectivo_mixto.set("")
        self.var_tipo_mixto.set("Efectivo + Transferencia")
        self._al_cambiar_tipo_mixto()
        if mp == "Mixto":
            self.var_metodo_pago.set("Transferencia")
        self._al_cambiar_metodo_pago()

        self.var_pago_vencido.set(False)
        self.var_editar_nro_recibo.set(False)
        self._toggle_editar_nro_recibo()
        self.var_editar_total_manual.set(False)
        self._toggle_total_manual()
        self.var_estado_pago.set("Pagado")
        self.var_monto_abonado.set("")
        self._al_cambiar_estado_pago()
        self.var_observaciones.set("")

        self._recargar_todo(inicial=False)
        self.root.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.yview_moveto(0.0)
        titular.autocomplete.entry.focus_set()

    def _refrescar_tabla(self, registros):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for r in registros:
            act_mostrar = r.get("Actividad", "")
            if act_mostrar == "Otro" and r.get("Detalle Actividad"):
                act_mostrar = f"Otro ({r.get('Detalle Actividad')})"
            elif act_mostrar == "Carnet" and r.get("Detalle Actividad"):
                act_mostrar = r.get("Detalle Actividad")

            carnet_mostrar = r.get("Nro Carnet", "")
            if isinstance(carnet_mostrar, float) and carnet_mostrar.is_integer():
                carnet_mostrar = int(carnet_mostrar)

            nro_rec = r.get("Nro Recibo", "")
            if isinstance(nro_rec, float) and nro_rec.is_integer():
                nro_rec = int(nro_rec)

            clases_s = r.get("Clases por Semana", "")
            if isinstance(clases_s, float) and clases_s.is_integer():
                clases_s = int(clases_s)

            estado_pago = str(r.get("Estado Pago", "Pagado")).strip()
            if not estado_pago or estado_pago.lower() == "nan":
                estado_pago = "Pagado"

            tag = ""
            if estado_pago == "Parcial":
                tag = "tag_parcial"
            elif estado_pago == "Debe":
                tag = "tag_debe"

            monto_tot = r.get("Total", "")
            if monto_tot is None or str(monto_tot).strip().lower() in ("", "nan", "none"):
                tot_mostrar = ""
            else:
                tot_mostrar = formatear_moneda(monto_tot)

            monto_pag = r.get("Monto Pagado", "")
            if monto_pag is None or str(monto_pag).strip().lower() in ("", "nan", "none"):
                pag_mostrar = ""
            else:
                pag_mostrar = formatear_moneda(monto_pag)

            monto_sal = r.get("Saldo Pendiente", "")
            if monto_sal is None or str(monto_sal).strip().lower() in ("", "nan", "none"):
                sal_mostrar = ""
            else:
                sal_mostrar = formatear_moneda(monto_sal)

            self.tree.insert(
                "",
                tk.END,
                values=(
                    nro_rec,
                    r.get("Fecha Emisión", ""),
                    r.get("Nombre", ""),
                    act_mostrar,
                    clases_s,
                    r.get("Fecha Vencimiento", ""),
                    r.get("Inscripción", ""),
                    r.get("Descuento", ""),
                    r.get("Método de Pago", ""),
                    carnet_mostrar,
                    tot_mostrar,
                    pag_mostrar,
                    sal_mostrar,
                    estado_pago,
                ),
                tags=(tag,) if tag else (),
            )

    def _eliminar_recibo_seleccionado(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleccionar", "Seleccioná primero una fila de la tabla de últimos recibos.")
            return

        valores = self.tree.item(sel[0], "values")
        nro_recibo = valores[0]
        nombre = valores[2]

        confirmar = messagebox.askyesno(
            "Confirmar Eliminación",
            f"¿Seguro que querés eliminar el Recibo N° {nro_recibo} (alumno/a '{nombre}') del archivo Excel?\n\n"
            "Si este recibo incluye más de una persona, se eliminarán todas las filas asociadas a ese N° de recibo.",
        )
        if not confirmar:
            return

        try:
            borrado = eliminar_recibo_por_numero(self.config["archivo_recibos"], nro_recibo)
            if borrado:
                self._recargar_todo(inicial=False)
                messagebox.showinfo("Eliminado", f"El Recibo N° {nro_recibo} fue eliminado correctamente.")
        except PermissionError:
            messagebox.showerror(
                "Archivo Abierto",
                "Cerrá 'recibos.xlsx' en Excel para poder eliminar el registro.",
            )

    def _abrir_dialogo_pago_pendiente(self, preseleccion=None):
        DialogoPagoPendiente(
            parent=self.root,
            ruta_recibos=self.config["archivo_recibos"],
            on_guardar_callback=lambda: self._recargar_todo(inicial=False),
            recibo_preseleccionado=preseleccion,
            precios_data=self.precios_data,
        )

    def _abrir_dialogo_pago_seleccionado(self):
        sel = self.tree.selection()
        if not sel:
            self._abrir_dialogo_pago_pendiente()
            return
        valores = self.tree.item(sel[0], "values")
        nro_recibo = valores[0]
        self._abrir_dialogo_pago_pendiente(preseleccion=nro_recibo)

    def _abrir_dialogo_carnet_sin_recibo(self):
        precio_carnet = self.precios_data.get("carnet", 1000.0)
        DialogoCarnetSinRecibo(
            parent=self.root,
            ruta_recibos=self.config["archivo_recibos"],
            lista_alumnos=self.lista_alumnos,
            precio_carnet=precio_carnet,
            proximo_carnet=self.proximo_carnet_auto,
            on_guardar_callback=lambda: self._recargar_todo(inicial=False),
        )


def main():
    root = tk.Tk()
    app = ReemplazoCuadernoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
