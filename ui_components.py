import os
import calendar
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date, datetime
from utils import quitar_tildes, formatear_moneda, parsear_fecha


# ==========================================================
# SELECTOR DE FECHA / CALENDARIO INTERACTIVO (POPUP)
# ==========================================================
class CalendarioPopup(tk.Toplevel):
    MESES = [
        "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    DIAS = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]

    def __init__(self, anchor_widget, var_fecha, permitir_vacio=False, on_select=None):
        super().__init__(anchor_widget.winfo_toplevel())
        self.anchor_widget = anchor_widget
        self.var_fecha = var_fecha
        self.permitir_vacio = permitir_vacio
        self.on_select = on_select

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg="#94a3b8", relief="solid", bd=1)

        val_str = self.var_fecha.get().strip()
        fecha_ini = date.today()
        self.dia_seleccionado = None
        self._anio_sel = None
        self._mes_sel = None

        if val_str:
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%Y-%m-%d"):
                try:
                    fecha_ini = datetime.strptime(val_str, fmt).date()
                    self.dia_seleccionado = fecha_ini.day
                    self._anio_sel = fecha_ini.year
                    self._mes_sel = fecha_ini.month
                    break
                except ValueError:
                    pass

        self.anio_actual = fecha_ini.year
        self.mes_actual = fecha_ini.month

        self._construir_ui()
        self._posicionar()

        self.bind("<Escape>", lambda e: self.destroy())
        self._root_bind_id = self.winfo_toplevel().bind_all("<ButtonPress-1>", self._al_click_global, add="+")
        self.bind("<Destroy>", self._al_destruir)

    def _al_destruir(self, event=None):
        if hasattr(self, "_root_bind_id") and self._root_bind_id:
            try:
                self.winfo_toplevel().unbind_all("<ButtonPress-1>")
            except Exception:
                pass
            self._root_bind_id = None
        if hasattr(self.anchor_widget, "_popup_calendario_activo"):
            self.anchor_widget._popup_calendario_activo = None

    def _al_click_global(self, event):
        if not self.winfo_exists():
            return
        try:
            x, y = self.winfo_rootx(), self.winfo_rooty()
            w, h = self.winfo_width(), self.winfo_height()
            if not (x <= event.x_root <= x + w and y <= event.y_root <= y + h):
                ax, ay = self.anchor_widget.winfo_rootx(), self.anchor_widget.winfo_rooty()
                aw, ah = self.anchor_widget.winfo_width(), self.anchor_widget.winfo_height()
                if not (ax <= event.x_root <= ax + aw and ay <= event.y_root <= ay + ah):
                    self.destroy()
        except Exception:
            self.destroy()

    def _posicionar(self):
        self.update_idletasks()
        try:
            x = self.anchor_widget.winfo_rootx()
            y = self.anchor_widget.winfo_rooty() + self.anchor_widget.winfo_height() + 2
            w = self.winfo_reqwidth()
            h = self.winfo_reqheight()

            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()

            if x + w > screen_w:
                x = screen_w - w - 10
            if y + h > screen_h:
                y = self.anchor_widget.winfo_rooty() - h - 2
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _construir_ui(self):
        contenedor = tk.Frame(self, bg="#ffffff", padx=8, pady=8)
        contenedor.pack(fill=tk.BOTH, expand=True)

        # Header de navegación
        f_nav = tk.Frame(contenedor, bg="#ffffff")
        f_nav.pack(fill=tk.X, pady=(0, 6))

        btn_anio_ant = tk.Button(
            f_nav, text="«", font=("Segoe UI", 9, "bold"), bg="#f1f5f9", fg="#334155",
            relief="flat", cursor="hand2", padx=4, pady=0, command=self._anio_ant
        )
        btn_anio_ant.pack(side=tk.LEFT, padx=(0, 2))

        btn_mes_ant = tk.Button(
            f_nav, text="‹", font=("Segoe UI", 9, "bold"), bg="#f1f5f9", fg="#334155",
            relief="flat", cursor="hand2", padx=5, pady=0, command=self._mes_ant
        )
        btn_mes_ant.pack(side=tk.LEFT, padx=(0, 4))

        self.lbl_mes_anio = tk.Label(
            f_nav, text="", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#0f172a"
        )
        self.lbl_mes_anio.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_cerrar = tk.Button(
            f_nav, text="✕", font=("Segoe UI", 8, "bold"), bg="#f1f5f9", fg="#64748b",
            relief="flat", cursor="hand2", padx=4, pady=0, command=self.destroy
        )
        btn_cerrar.pack(side=tk.RIGHT, padx=(4, 0))

        btn_anio_sig = tk.Button(
            f_nav, text="»", font=("Segoe UI", 9, "bold"), bg="#f1f5f9", fg="#334155",
            relief="flat", cursor="hand2", padx=4, pady=0, command=self._anio_sig
        )
        btn_anio_sig.pack(side=tk.RIGHT, padx=(2, 0))

        btn_mes_sig = tk.Button(
            f_nav, text="›", font=("Segoe UI", 9, "bold"), bg="#f1f5f9", fg="#334155",
            relief="flat", cursor="hand2", padx=5, pady=0, command=self._mes_sig
        )
        btn_mes_sig.pack(side=tk.RIGHT, padx=(2, 0))

        # Días semana
        f_dias = tk.Frame(contenedor, bg="#ffffff")
        f_dias.pack(fill=tk.X, pady=(2, 2))
        for d in self.DIAS:
            tk.Label(
                f_dias, text=d, font=("Segoe UI", 8, "bold"), width=3,
                bg="#ffffff", fg="#64748b"
            ).pack(side=tk.LEFT, padx=1)

        # Grilla de días
        self.f_grid = tk.Frame(contenedor, bg="#ffffff")
        self.f_grid.pack(fill=tk.X, pady=(2, 6))

        # Barra inferior
        f_inferior = tk.Frame(contenedor, bg="#ffffff")
        f_inferior.pack(fill=tk.X)

        btn_hoy = tk.Button(
            f_inferior, text="📅 Hoy", font=("Segoe UI", 8, "bold"),
            bg="#e0f2fe", fg="#0369a1", relief="flat", cursor="hand2",
            padx=6, pady=2, command=self._seleccionar_hoy
        )
        btn_hoy.pack(side=tk.LEFT, padx=(0, 4))

        if self.permitir_vacio:
            btn_vacio = tk.Button(
                f_inferior, text="❌ Dejar vacío", font=("Segoe UI", 8, "bold"),
                bg="#fee2e2", fg="#991b1b", relief="flat", cursor="hand2",
                padx=6, pady=2, command=self._limpiar_fecha
            )
            btn_vacio.pack(side=tk.LEFT)

        self._renderizar_mes()

    def _renderizar_mes(self):
        for w in self.f_grid.winfo_children():
            w.destroy()

        self.lbl_mes_anio.config(text=f"{self.MESES[self.mes_actual]} {self.anio_actual}")
        cal = calendar.monthcalendar(self.anio_actual, self.mes_actual)
        hoy = date.today()

        for sem in cal:
            f_sem = tk.Frame(self.f_grid, bg="#ffffff")
            f_sem.pack(fill=tk.X, pady=1)
            for d in sem:
                if d == 0:
                    tk.Label(f_sem, text="", width=3, bg="#ffffff").pack(side=tk.LEFT, padx=1)
                else:
                    es_hoy = (self.anio_actual == hoy.year and self.mes_actual == hoy.month and d == hoy.day)
                    es_sel = (
                        self.dia_seleccionado == d and
                        self.anio_actual == self._anio_sel and
                        self.mes_actual == self._mes_sel
                    )

                    bg_color = "#2563eb" if es_sel else ("#e0f2fe" if es_hoy else "#f8fafc")
                    fg_color = "#ffffff" if es_sel else ("#0284c7" if es_hoy else "#1e293b")
                    font_d = ("Segoe UI", 8, "bold" if (es_hoy or es_sel) else "normal")

                    btn = tk.Button(
                        f_sem, text=str(d), width=3, font=font_d,
                        bg=bg_color, fg=fg_color, relief="flat", cursor="hand2",
                        command=lambda dia=d: self._seleccionar_dia(dia)
                    )
                    btn.pack(side=tk.LEFT, padx=1)

    def _mes_ant(self):
        if self.mes_actual == 1:
            self.mes_actual = 12
            self.anio_actual -= 1
        else:
            self.mes_actual -= 1
        self._renderizar_mes()

    def _mes_sig(self):
        if self.mes_actual == 12:
            self.mes_actual = 1
            self.anio_actual += 1
        else:
            self.mes_actual += 1
        self._renderizar_mes()

    def _anio_ant(self):
        self.anio_actual -= 1
        self._renderizar_mes()

    def _anio_sig(self):
        self.anio_actual += 1
        self._renderizar_mes()

    def _seleccionar_dia(self, dia):
        f = date(self.anio_actual, self.mes_actual, dia)
        f_str = f.strftime("%d/%m/%Y")
        self.var_fecha.set(f_str)
        if self.on_select:
            self.on_select(f_str)
        self.destroy()

    def _seleccionar_hoy(self):
        hoy_str = date.today().strftime("%d/%m/%Y")
        self.var_fecha.set(hoy_str)
        if self.on_select:
            self.on_select(hoy_str)
        self.destroy()

    def _limpiar_fecha(self):
        self.var_fecha.set("")
        if self.on_select:
            self.on_select("")
        self.destroy()


def abrir_calendario(anchor_widget, var_fecha, permitir_vacio=False, on_select=None):
    """Abre el popup interactivo de calendario anclado a un botón o entry."""
    if hasattr(anchor_widget, "_popup_calendario_activo") and anchor_widget._popup_calendario_activo:
        try:
            if anchor_widget._popup_calendario_activo.winfo_exists():
                anchor_widget._popup_calendario_activo.destroy()
                anchor_widget._popup_calendario_activo = None
                return
        except Exception:
            pass

    pop = CalendarioPopup(anchor_widget, var_fecha, permitir_vacio=permitir_vacio, on_select=on_select)
    anchor_widget._popup_calendario_activo = pop
    return pop


# ==========================================================
# COMPONENTE DE AUTOCOMPLETADO PARA NOMBRE DE ALUMNOS
# ==========================================================
class AutocompleteEntry(ttk.Frame):
    def __init__(self, parent, textvariable, on_select_callback=None, **kwargs):
        super().__init__(parent)
        self.textvariable = textvariable
        self.on_select_callback = on_select_callback
        self.lista_completa = []
        self._bloqueando_evento = False

        self.entry = ttk.Entry(self, textvariable=self.textvariable, font=("Segoe UI", 10), **kwargs)
        self.entry.pack(fill=tk.X, expand=True)

        self.popup = None
        self.listbox = None

        self.entry.bind("<KeyRelease>", self._al_escribir)
        self.entry.bind("<Down>", self._foco_lista)
        self.entry.bind("<Escape>", lambda e: self.cerrar_popup())
        self.entry.bind("<Return>", self._seleccionar_primero_si_abierto)
        self.entry.bind("<FocusOut>", self._al_perder_foco)

    def actualizar_base(self, nombres):
        self.lista_completa = list(nombres)

    def _al_escribir(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return
        if self._bloqueando_evento:
            return

        texto = self.textvariable.get().strip()
        if not texto or not self.lista_completa:
            self.cerrar_popup()
            return

        buscado = quitar_tildes(texto)
        palabras = [p for p in buscado.split() if p]

        empiezan = []
        contienen = []
        for item in self.lista_completa:
            item_norm = quitar_tildes(item)
            if item_norm.startswith(buscado):
                empiezan.append(item)
            elif all(p in item_norm for p in palabras):
                contienen.append(item)

        coincidencias = (empiezan + contienen)[:12]
        if coincidencias:
            self._mostrar_popup(coincidencias)
        else:
            self.cerrar_popup()

    def _mostrar_popup(self, coincidencias):
        if not self.popup or not self.popup.winfo_exists():
            self.popup = tk.Toplevel(self)
            self.popup.wm_overrideredirect(True)
            self.popup.attributes("-topmost", True)

            frame_list = ttk.Frame(self.popup, borderwidth=1, relief="solid")
            frame_list.pack(fill=tk.BOTH, expand=True)

            self.listbox = tk.Listbox(
                frame_list,
                font=("Segoe UI", 10),
                height=min(len(coincidencias), 7),
                activestyle="none",
                selectbackground="#0284c7",
                selectforeground="white",
            )
            scroll = ttk.Scrollbar(frame_list, orient="vertical", command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=scroll.set)

            self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll.pack(side=tk.RIGHT, fill=tk.Y)

            self.listbox.bind("<ButtonRelease-1>", self._confirmar_seleccion)
            self.listbox.bind("<Return>", self._confirmar_seleccion)
            self.listbox.bind("<Escape>", lambda e: self.cerrar_popup())
            self.listbox.bind("<Up>", self._manejar_arriba_listbox)
            self.listbox.bind("<FocusOut>", self._al_perder_foco)

        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        w = max(self.entry.winfo_width(), 260)
        h = min(len(coincidencias) * 23 + 6, 170)
        self.popup.geometry(f"{w}x{h}+{x}+{y}")

        self.listbox.delete(0, tk.END)
        for nombre in coincidencias:
            self.listbox.insert(tk.END, f"  {nombre}")
        self.listbox.selection_set(0)

    def _foco_lista(self, event):
        if self.popup and self.popup.winfo_exists() and self.listbox and self.listbox.size() > 0:
            self.listbox.focus_set()
            if not self.listbox.curselection():
                self.listbox.selection_set(0)
            return "break"

    def _manejar_arriba_listbox(self, event):
        sel = self.listbox.curselection()
        if sel and sel[0] == 0:
            self.entry.focus_set()
            return "break"

    def _seleccionar_primero_si_abierto(self, event):
        if self.popup and self.popup.winfo_exists() and self.listbox and self.listbox.size() > 0:
            self._confirmar_seleccion()
            return "break"

    def _confirmar_seleccion(self, event=None):
        if self.listbox and self.listbox.curselection():
            sel_idx = self.listbox.curselection()[0]
            val = self.listbox.get(sel_idx).strip()
            self._bloqueando_evento = True
            self.textvariable.set(val)
            self._bloqueando_evento = False
            self.cerrar_popup()
            self.entry.icursor(tk.END)
            if self.on_select_callback:
                self.on_select_callback(val)
        return "break"

    def _al_perder_foco(self, event):
        self.after(150, self._verificar_foco)

    def _verificar_foco(self):
        foco_actual = self.focus_get()
        if self.popup and self.popup.winfo_exists():
            if foco_actual not in (self.entry, self.listbox):
                self.cerrar_popup()

    def cerrar_popup(self):
        if self.popup and self.popup.winfo_exists():
            self.popup.destroy()
        self.popup = None
        self.listbox = None


# ==========================================================
# COMPONENTE VISUAL PARA CADA ALUMNO DENTRO DEL RECIBO
# ==========================================================
class AlumnoItem:
    def __init__(self, parent_frame, index, app, es_titular=True):
        self.parent_frame = parent_frame
        self.index = index
        self.app = app
        self.es_titular = es_titular

        self.var_nombre = tk.StringVar()
        self.var_usar_carnet = tk.BooleanVar(value=False)
        self.var_nro_carnet = tk.StringVar(value="")
        self.var_actividad = tk.StringVar(value="Natación")
        self.var_actividad_otro = tk.StringVar()
        self.var_clases_sem = tk.IntVar(value=2)
        self.var_inscripcion = tk.StringVar(value="No")
        self.var_gf_asociado = tk.StringVar(value="")
        self.var_fecha_vencimiento = tk.StringVar(
            value=self.app.var_fecha_vencimiento.get() if hasattr(self.app, "var_fecha_vencimiento") else ""
        )

        # Regla: Si es el segundo o posterior, marcar Grupo familiar por defecto automáticamente!
        if self.es_titular:
            mp = self.app.var_metodo_pago.get() if hasattr(self.app, "var_metodo_pago") else "Transferencia"
            self.var_descuento = tk.StringVar(value="Efectivo" if mp == "Efectivo" else "Ninguno")
        else:
            self.var_descuento = tk.StringVar(value="Grupo familiar")

        self.frame = None
        self.lbl_subtotal = None
        self.lbl_dias_clase = None
        self.autocomplete = None
        self.entry_carnet = None
        self.chk_carnet = None
        self.entry_venc = None
        self.entry_actividad_otro = None
        self.f_gf_asociado = None
        self.entry_gf_asociado = None

        self._construir_ui()
        self._vincular_trazas()

    def _construir_ui(self):
        tit_txt = f" Alumno {self.index} (Titular) " if self.es_titular else f" Alumno {self.index} (Grupo Familiar - GF) "
        self.frame = ttk.LabelFrame(self.parent_frame, text=tit_txt, padding="10 5 10 5")
        self.frame.pack(fill=tk.X, pady=(0, 6))

        # --- Fila 1: Nombre, Días de clase, Vencimiento, Carnet opcional, Botón Quitar ---
        f1 = ttk.Frame(self.frame)
        f1.pack(fill=tk.X, pady=(0, 4))

        f_nom = ttk.Frame(f1)
        f_nom.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))
        lbl_nom_txt = f"Nombre del Alumno {self.index} (autocompletado):"
        ttk.Label(f_nom, text=lbl_nom_txt, style="Field.TLabel").pack(anchor=tk.W)
        self.autocomplete = AutocompleteEntry(f_nom, textvariable=self.var_nombre, on_select_callback=self._al_seleccionar_nombre)
        self.autocomplete.actualizar_base(self.app.lista_alumnos)
        self.autocomplete.pack(fill=tk.X, pady=(2, 0))

        # Etiqueta para mostrar los días de clase extraídos
        self.lbl_dias_clase = ttk.Label(f_nom, text="", font=("Segoe UI", 8, "italic"), foreground="#0284c7")
        self.lbl_dias_clase.pack(anchor=tk.W, pady=(2, 0))

        # Vencimiento individual por alumno
        self.f_venc = ttk.Frame(f1)
        self.f_venc.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(self.f_venc, text="Vencimiento:", style="Field.TLabel").pack(anchor=tk.W)
        f_ven_in = ttk.Frame(self.f_venc)
        f_ven_in.pack(anchor=tk.W, pady=(2, 0))
        self.entry_venc = ttk.Entry(
            f_ven_in,
            textvariable=self.var_fecha_vencimiento,
            width=11,
            font=("Segoe UI", 9),
            justify="center",
        )
        self.entry_venc.pack(side=tk.LEFT, padx=(0, 2))
        self.entry_venc.bind("<Double-Button-1>", lambda e: self._abrir_cal_venc())
        self.btn_cal_venc = ttk.Button(
            f_ven_in,
            text="📅",
            width=3,
            command=self._abrir_cal_venc,
        )
        self.btn_cal_venc.pack(side=tk.LEFT)

        f_car = ttk.Frame(f1)
        f_car.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(f_car, text="Carnet (Opcional):", style="Field.TLabel").pack(anchor=tk.W)
        f_car_in = ttk.Frame(f_car)
        f_car_in.pack(anchor=tk.W, pady=(2, 0))
        self.chk_carnet = ttk.Checkbutton(
            f_car_in,
            text="Asignar N°:",
            variable=self.var_usar_carnet,
            command=self._toggle_carnet,
        )
        self.chk_carnet.pack(side=tk.LEFT, padx=(0, 4))
        self.entry_carnet = ttk.Entry(
            f_car_in,
            textvariable=self.var_nro_carnet,
            width=8,
            font=("Segoe UI", 10, "bold"),
            justify="center",
            state="disabled",
        )
        self.entry_carnet.pack(side=tk.LEFT)

        if not self.es_titular:
            f_btn_del = ttk.Frame(f1)
            f_btn_del.pack(side=tk.RIGHT, pady=(12, 0))
            ttk.Button(
                f_btn_del,
                text="❌ Quitar",
                width=9,
                command=lambda: self.app._quitar_alumno(self),
            ).pack(side=tk.RIGHT)

        # --- Fila 2: Actividad y Clases/semana ---
        f2 = ttk.Frame(self.frame)
        f2.pack(fill=tk.X, pady=(2, 3))

        # Actividad
        f_act = ttk.Frame(f2)
        f_act.pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(f_act, text="Actividad:", style="Field.TLabel").pack(anchor=tk.W)
        self.f_act_opts = ttk.Frame(f_act)
        self.f_act_opts.pack(anchor=tk.W, pady=(2, 0))
        for act_op in ("Natación", "Aquagym", "No figura", "Clase Adaptación (CdA)", "Solo Inscripción", "Otro"):
            ttk.Radiobutton(
                self.f_act_opts,
                text=act_op,
                value=act_op,
                variable=self.var_actividad,
                command=self._al_cambiar_actividad,
            ).pack(side=tk.LEFT, padx=(0, 6))
        self.entry_actividad_otro = ttk.Entry(
            self.f_act_opts,
            textvariable=self.var_actividad_otro,
            width=12,
            font=("Segoe UI", 9),
            state="disabled",
        )
        self.entry_actividad_otro.pack(side=tk.LEFT)

        # Clases/sem
        self.f_cla = ttk.Frame(f2)
        self.f_cla.pack(side=tk.LEFT)
        self.lbl_cla_titulo = ttk.Label(self.f_cla, text="Clases/sem:", style="Field.TLabel")
        self.lbl_cla_titulo.pack(anchor=tk.W)
        self.f_cla_opts = ttk.Frame(self.f_cla)
        self.f_cla_opts.pack(anchor=tk.W, pady=(2, 0))
        for cant, etiq in ((1, "1 (4/mes)"), (2, "2 (8/mes)"), (3, "3 (12/mes)")):
            ttk.Radiobutton(
                self.f_cla_opts,
                text=etiq,
                value=cant,
                variable=self.var_clases_sem,
            ).pack(side=tk.LEFT, padx=(0, 6))

        # --- Fila 3: Inscripción y Descuento ---
        f3 = ttk.Frame(self.frame)
        f3.pack(fill=tk.X, pady=(2, 3))

        # Inscripción
        self.f_ins = ttk.Frame(f3)
        self.f_ins.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(self.f_ins, text="Inscripción:", style="Field.TLabel").pack(anchor=tk.W)
        self.f_ins_opts = ttk.Frame(self.f_ins)
        self.f_ins_opts.pack(anchor=tk.W, pady=(2, 0))
        ttk.Radiobutton(self.f_ins_opts, text="No ($0)", value="No", variable=self.var_inscripcion).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Radiobutton(self.f_ins_opts, text="Sí", value="Sí", variable=self.var_inscripcion).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Radiobutton(self.f_ins_opts, text="CdA ($0)", value="CdA", variable=self.var_inscripcion).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Radiobutton(self.f_ins_opts, text="Ticket Sochi", value="Ticket Sochi", variable=self.var_inscripcion).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Radiobutton(self.f_ins_opts, text="Ticket Malva", value="Ticket Malva", variable=self.var_inscripcion).pack(side=tk.LEFT)

        # Descuento
        self.f_des = ttk.Frame(f3)
        self.f_des.pack(side=tk.LEFT)
        ttk.Label(self.f_des, text="Descuento en Cuota:", style="Field.TLabel").pack(anchor=tk.W)
        self.f_des_opts = ttk.Frame(self.f_des)
        self.f_des_opts.pack(anchor=tk.W, pady=(2, 0))
        for d_op in ("Ninguno", "Efectivo", "Grupo familiar", "Jubilado"):
            ttk.Radiobutton(
                self.f_des_opts,
                text=d_op,
                value=d_op,
                variable=self.var_descuento,
            ).pack(side=tk.LEFT, padx=(0, 6))

        # Detalle de familiar para quien paga solo con descuento de Grupo Familiar
        self.f_gf_asociado = ttk.Frame(f3)
        ttk.Label(self.f_gf_asociado, text="Familiar / Recibo GF:", font=("Segoe UI", 9, "bold"), foreground="#1e40af").pack(
            side=tk.LEFT, padx=(8, 4)
        )
        self.entry_gf_asociado = ttk.Entry(
            self.f_gf_asociado, textvariable=self.var_gf_asociado, width=22, font=("Segoe UI", 9)
        )
        self.entry_gf_asociado.pack(side=tk.LEFT)

        # --- Fila 4: Subtotal individual del alumno ---
        f4 = ttk.Frame(self.frame)
        f4.pack(fill=tk.X, pady=(2, 0))
        self.lbl_subtotal = ttk.Label(
            f4,
            text=f"Subtotal Alumno {self.index}: $0",
            font=("Segoe UI", 9, "bold"),
            foreground="#1e3a8a",
        )
        self.lbl_subtotal.pack(side=tk.LEFT)

    def _vincular_trazas(self):
        cb = lambda *args: self.app._actualizar_calculo_precio()
        self.var_clases_sem.trace_add("write", cb)
        self.var_inscripcion.trace_add("write", cb)
        self.var_descuento.trace_add("write", self._al_cambiar_descuento)
        self.var_gf_asociado.trace_add("write", cb)
        self.var_actividad.trace_add("write", cb)
        self.var_nombre.trace_add("write", self._al_cambiar_nombre)
        self.var_fecha_vencimiento.trace_add("write", cb)

    def _al_seleccionar_nombre(self, nombre_elegido):
        self._al_cambiar_nombre()

    def _al_cambiar_nombre(self, *args):
        nom = self.var_nombre.get().strip()
        if not nom:
            if self.lbl_dias_clase:
                self.lbl_dias_clase.config(text="")
            self.app._actualizar_calculo_precio()
            return

        norm = quitar_tildes(nom).lower().strip()
        detalles = getattr(self.app, "detalles_alumnos", {})
        det = detalles.get(norm) or detalles.get(nom)

        if det:
            dias = det.get("dias_clase", "")
            if dias:
                txt_dias = f"📅 Días de clase: {dias}"
            elif det.get("horario_raw", ""):
                txt_dias = f"📅 Horario registrado: {det.get('horario_raw')}"
            else:
                txt_dias = "📅 Días de clase: Sin registrar"
            if self.lbl_dias_clase:
                self.lbl_dias_clase.config(text=txt_dias)

            # Si la 3ra columna dice jubilado/a, marcar directamente la opción Jubilado!
            if det.get("es_jubilado"):
                self.var_descuento.set("Jubilado")
        else:
            if self.lbl_dias_clase:
                self.lbl_dias_clase.config(text="")

        self.app._actualizar_calculo_precio()

    def _al_cambiar_descuento(self, *args):
        self._actualizar_visibilidad_gf_asociado()
        self.app._actualizar_calculo_precio()

    def _actualizar_visibilidad_gf_asociado(self):
        if hasattr(self, "f_gf_asociado") and self.f_gf_asociado:
            es_solo = len(getattr(self.app, "alumnos_items", [])) <= 1
            if self.var_descuento.get() == "Grupo familiar" and (self.es_titular or es_solo):
                self.f_gf_asociado.pack(side=tk.LEFT, padx=(10, 0))
            else:
                self.f_gf_asociado.pack_forget()

    def actualizar_indice(self, nuevo_idx):
        self.index = nuevo_idx
        tit_txt = f" Alumno {self.index} (Titular) " if self.es_titular else f" Alumno {self.index} (Grupo Familiar - GF) "
        if self.frame:
            self.frame.config(text=tit_txt)
        self._actualizar_visibilidad_gf_asociado()
        self.actualizar_etiqueta_subtotal()

    def _toggle_carnet(self):
        if self.var_usar_carnet.get():
            self.entry_carnet.config(state="normal")
            if not self.var_nro_carnet.get().strip():
                sug = self.app._sugerir_proximo_carnet(self)
                self.var_nro_carnet.set(str(sug))
            self.entry_carnet.focus_set()
        else:
            self.entry_carnet.config(state="disabled")
            self.var_nro_carnet.set("")

    def _al_cambiar_actividad(self):
        act = self.var_actividad.get()
        if act == "Otro":
            self.entry_actividad_otro.config(state="normal")
            self.entry_actividad_otro.focus_set()
        else:
            self.var_actividad_otro.set("")
            self.entry_actividad_otro.config(state="disabled")

        if act == "Clase Adaptación (CdA)":
            self._habilitar_clases(False)
            self._habilitar_inscripcion(False)
            self._habilitar_descuento(False)
            self.var_inscripcion.set("No")
            self.var_fecha_vencimiento.set("")
            if hasattr(self, "entry_venc") and self.entry_venc:
                self.entry_venc.config(state="disabled")
            if hasattr(self, "btn_cal_venc") and self.btn_cal_venc:
                self.btn_cal_venc.config(state="disabled")
        elif act == "Solo Inscripción":
            self._habilitar_clases(False)
            self._habilitar_inscripcion(True)
            if self.var_inscripcion.get() == "No":
                self.var_inscripcion.set("Sí")
            self._habilitar_descuento(False)
            if hasattr(self, "entry_venc") and self.entry_venc:
                self.entry_venc.config(state="normal")
            if hasattr(self, "btn_cal_venc") and self.btn_cal_venc:
                self.btn_cal_venc.config(state="normal")
        else:
            # Natación, Aquagym, No figura, Otro
            self._habilitar_clases(True)
            self._habilitar_inscripcion(True)
            self._habilitar_descuento(True)
            if hasattr(self, "entry_venc") and self.entry_venc:
                self.entry_venc.config(state="normal")
            if hasattr(self, "btn_cal_venc") and self.btn_cal_venc:
                self.btn_cal_venc.config(state="normal")
            if not self.var_fecha_vencimiento.get().strip() and hasattr(self.app, "var_fecha_vencimiento"):
                self.var_fecha_vencimiento.set(self.app.var_fecha_vencimiento.get())

        self.app._actualizar_calculo_precio()

    def _abrir_cal_venc(self):
        if self.var_actividad.get() == "Clase Adaptación (CdA)":
            return
        abrir_calendario(self.btn_cal_venc, self.var_fecha_vencimiento, permitir_vacio=True)

    def _toggle_actividad_otro(self):
        self._al_cambiar_actividad()

    def _habilitar_clases(self, habilitar):
        st = "normal" if habilitar else "disabled"
        for w in self.f_cla_opts.winfo_children():
            w.config(state=st)

    def _habilitar_inscripcion(self, habilitar):
        st = "normal" if habilitar else "disabled"
        for w in self.f_ins_opts.winfo_children():
            w.config(state=st)

    def _habilitar_descuento(self, habilitar):
        st = "normal" if habilitar else "disabled"
        for w in self.f_des_opts.winfo_children():
            w.config(state=st)

    def obtener_montos(self):
        actividad = self.var_actividad.get()
        precio_inscripcion_base = float(self.app.precios_data.get("inscripcion", 15000.0))

        if actividad == "Clase Adaptación (CdA)":
            monto_cuota = precio_inscripcion_base  # Mismo valor que la inscripción
            monto_insc = 0.0
            return monto_cuota, monto_insc, monto_cuota

        if actividad == "Solo Inscripción":
            monto_cuota = 0.0
            tipo_insc = self.var_inscripcion.get()
            monto_insc = precio_inscripcion_base if tipo_insc == "Sí" else 0.0
            return monto_cuota, monto_insc, monto_insc

        # Actividad regular (Natación, Aquagym, No figura, Otro)
        clases = self.var_clases_sem.get()
        cuotas_dict = self.app.precios_data.get("cuotas", {}).get(clases, {})

        fuera_termino = hasattr(self.app, "var_pago_vencido") and self.app.var_pago_vencido.get()
        desc_sel = self.var_descuento.get()

        if fuera_termino:
            descuento = "Ninguno"
        elif desc_sel == "Jubilado":
            # Regla: Pago mixto (o cualquier método) para jubilados tiene descuento jub independiente de la cantidad de efectivo
            descuento = "Jubilado"
        elif desc_sel == "Grupo familiar":
            # Permite descuento de Grupo Familiar tanto en grupo como individual
            descuento = "Grupo familiar"
        elif self.es_titular:
            mp = self.app.var_metodo_pago.get() if hasattr(self.app, "var_metodo_pago") else "Transferencia"
            if mp == "Efectivo":
                descuento = "Efectivo"
            elif mp == "Mixto":
                tipo_m = self.app.var_tipo_mixto.get() if hasattr(self.app, "var_tipo_mixto") else "Efectivo + Transferencia"
                if "Efectivo" in tipo_m:
                    cuota_ef_titular = float(cuotas_dict.get("Efectivo", 0.0))
                    monto_ef = self.app._obtener_efectivo_mixto() if hasattr(self.app, "_obtener_efectivo_mixto") else 0.0
                    if monto_ef >= cuota_ef_titular and cuota_ef_titular > 0:
                        descuento = "Efectivo"
                    else:
                        descuento = "Ninguno"
                else:
                    # Transferencia + Débito no tiene efectivo, no aplica descuento de efectivo
                    descuento = "Ninguno"
            else:
                descuento = "Ninguno"
        else:
            descuento = desc_sel

        monto_cuota = float(cuotas_dict.get(descuento, cuotas_dict.get("Ninguno", 0.0)))

        tipo_insc = self.var_inscripcion.get()
        if tipo_insc == "Sí":
            monto_insc = precio_inscripcion_base
        else:
            # "No", "CdA", "Ticket Sochi", "Ticket Malva"
            monto_insc = 0.0

        subtotal = monto_cuota + monto_insc
        return monto_cuota, monto_insc, subtotal

    def actualizar_etiqueta_subtotal(self):
        cuota, insc, subtotal = self.obtener_montos()
        actividad = self.var_actividad.get()

        if actividad == "Clase Adaptación (CdA)":
            txt = f"Subtotal Alumno {self.index}: {formatear_moneda(subtotal)}  (Clase de Adaptación CdA: {formatear_moneda(cuota)})"
        elif actividad == "Solo Inscripción":
            tipo_insc = self.var_inscripcion.get()
            if tipo_insc in ("Ticket Sochi", "Ticket Malva", "Sochi", "Malva"):
                txt = f"Subtotal Alumno {self.index}: {formatear_moneda(subtotal)}  (Solo Inscripción: $0 - {tipo_insc})"
            elif tipo_insc == "CdA":
                txt = f"Subtotal Alumno {self.index}: {formatear_moneda(subtotal)}  (Solo Inscripción: $0 - CdA ya pagada)"
            else:
                txt = f"Subtotal Alumno {self.index}: {formatear_moneda(subtotal)}  (Solo Inscripción: {formatear_moneda(insc)})"
        else:
            fuera_term = hasattr(self.app, "var_pago_vencido") and self.app.var_pago_vencido.get()
            desc_sel = self.var_descuento.get()
            if fuera_term:
                desc_txt = "Fuera de término (Sin dto.)"
            elif desc_sel == "Jubilado":
                desc_txt = "Jubilado/a"
            elif desc_sel == "Grupo familiar" and self.es_titular:
                gf_extra = f" [{self.var_gf_asociado.get().strip()}]" if self.var_gf_asociado.get().strip() else ""
                desc_txt = f"Grupo familiar{gf_extra}"
            elif self.es_titular and hasattr(self.app, "var_metodo_pago") and self.app.var_metodo_pago.get() == "Mixto":
                tipo_m = self.app.var_tipo_mixto.get() if hasattr(self.app, "var_tipo_mixto") else "Efectivo + Transferencia"
                if "Efectivo" in tipo_m:
                    cuotas_dict = self.app.precios_data.get("cuotas", {}).get(self.var_clases_sem.get(), {})
                    cuota_ef_titular = float(cuotas_dict.get("Efectivo", 0.0))
                    monto_ef = self.app._obtener_efectivo_mixto() if hasattr(self.app, "_obtener_efectivo_mixto") else 0.0
                    if monto_ef >= cuota_ef_titular and cuota_ef_titular > 0:
                        desc_txt = "Efectivo (Mixto: cubre cuota titular)"
                    else:
                        desc_txt = "Ninguno (Mixto: ef. no cubre cuota titular)"
                else:
                    desc_txt = "Ninguno (Mixto: Tr + Deb sin efectivo)"
            else:
                desc_txt = desc_sel

            tipo_insc = self.var_inscripcion.get()
            if tipo_insc in ("Ticket Sochi", "Ticket Malva", "Sochi", "Malva"):
                insc_txt = f"Inscripción: $0 ({tipo_insc})"
            elif tipo_insc == "CdA":
                insc_txt = "Inscripción: $0 (CdA ya pagada)"
            else:
                insc_txt = f"Inscripción: {formatear_moneda(insc)}"

            txt = (
                f"Subtotal Alumno {self.index}: {formatear_moneda(subtotal)}  "
                f"(Cuota {self.var_clases_sem.get() * 4} clases/mes [{desc_txt}]: {formatear_moneda(cuota)} | "
                f"{insc_txt})"
            )

        if self.lbl_subtotal:
            self.lbl_subtotal.config(text=txt)


# ==========================================================
# DIÁLOGO MODAL PARA CONFIGURACIÓN DE BASE DE ALUMNOS
# ==========================================================
class DialogoConfigAlumnos:
    def __init__(self, parent, config, on_guardar_callback):
        self.dlg = tk.Toplevel(parent)
        self.dlg.title("Configurar Base de Alumnos (Link de OneDrive)")
        self.dlg.geometry("640x360")
        self.dlg.resizable(False, False)
        self.dlg.transient(parent)
        self.dlg.grab_set()

        self.config = config
        self.on_guardar_callback = on_guardar_callback

        frame = ttk.Frame(self.dlg, padding="16 14 16 14")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame,
            text="Conexión con Base de Datos de Alumnos (Solo lectura - 2da columna, 1ra hoja)",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor=tk.W, pady=(0, 8))

        # Opción 1: Link compartido
        lf_link = ttk.LabelFrame(frame, text=" Opción 1: Pegar Link compartido de OneDrive / Google Sheets ", padding="10 8 10 8")
        lf_link.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            lf_link,
            text="Pegá acá el link del archivo compartido (ej. https://1drv.ms/... o https://...-my.sharepoint.com/...):",
        ).pack(anchor=tk.W, pady=(0, 4))

        self.var_url = tk.StringVar(value=self.config.get("url_onedrive_alumnos", ""))
        entry_url = ttk.Entry(lf_link, textvariable=self.var_url, font=("Segoe UI", 9))
        entry_url.pack(fill=tk.X, pady=(0, 6))

        f_freq = ttk.Frame(lf_link)
        f_freq.pack(fill=tk.X)
        ttk.Label(f_freq, text="Actualizar automáticamente:").pack(side=tk.LEFT, padx=(0, 8))
        self.var_freq = tk.StringVar(value=self.config.get("frecuencia_sync_alumnos", "diaria"))
        ttk.Radiobutton(f_freq, text="1 vez por día (al abrir el programa)", value="diaria", variable=self.var_freq).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Radiobutton(f_freq, text="Cada vez que se abre el programa", value="siempre", variable=self.var_freq).pack(
            side=tk.LEFT
        )

        # Opción 2: Archivo local
        lf_local = ttk.LabelFrame(
            frame,
            text=" Opción 2 (Alternativa): Archivo sincronizado en la PC ",
            padding="10 8 10 8",
        )
        lf_local.pack(fill=tk.X, pady=(0, 12))

        self.var_local = tk.StringVar(value=self.config.get("archivo_alumnos", ""))
        f_loc_row = ttk.Frame(lf_local)
        f_loc_row.pack(fill=tk.X)
        ttk.Entry(f_loc_row, textvariable=self.var_local, font=("Segoe UI", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6)
        )

        def elegir_local():
            r = filedialog.askopenfilename(
                parent=self.dlg,
                title="Seleccionar archivo Excel",
                filetypes=[("Archivos Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")],
            )
            if r:
                self.var_local.set(os.path.abspath(r))

        ttk.Button(f_loc_row, text="Examinar...", command=elegir_local).pack(side=tk.RIGHT)

        ttk.Label(
            lf_local,
            text="💡 Tip: Si en OneDrive Web abrís 'Compartido' y tocás 'Agregar acceso directo a Mis archivos', aparece en tu PC.",
            style="Small.TLabel",
        ).pack(anchor=tk.W, pady=(4, 0))

        # Botones guardar
        f_btns = ttk.Frame(frame)
        f_btns.pack(fill=tk.X, pady=(4, 0))

        ttk.Button(
            f_btns,
            text="💾 Guardar y Sincronizar Ahora",
            style="Save.TButton",
            command=self._guardar,
        ).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(f_btns, text="Cancelar", command=self.dlg.destroy).pack(side=tk.RIGHT)

    def _guardar(self):
        self.config["url_onedrive_alumnos"] = self.var_url.get().strip()
        self.config["archivo_alumnos"] = self.var_local.get().strip()
        self.config["frecuencia_sync_alumnos"] = self.var_freq.get()
        self.dlg.destroy()
        if self.on_guardar_callback:
            self.on_guardar_callback()


# ==========================================================
# DIÁLOGO MODAL PARA COBRO DE SALDOS PENDIENTES
# ==========================================================
class DialogoPagoPendiente:
    def __init__(self, parent, ruta_recibos, on_guardar_callback, recibo_preseleccionado=None, precios_data=None):
        self.parent = parent
        self.ruta_recibos = ruta_recibos
        self.on_guardar_callback = on_guardar_callback
        self.precios_data = precios_data or {}

        from excel_manager import obtener_recibos_con_deuda, registrar_pago_pendiente
        self._registrar_fn = registrar_pago_pendiente
        self.recibos_con_deuda = obtener_recibos_con_deuda(self.ruta_recibos)

        if not self.recibos_con_deuda:
            messagebox.showinfo(
                "Sin Deudas Pendientes",
                "¡Excelente! No hay recibos con saldo pendiente o pagos a medias registrados en el Excel.",
                parent=self.parent
            )
            return

        self.dlg = tk.Toplevel(parent)
        self.dlg.title("Registrar Pago Pendiente / Cobro de Deuda")
        self.dlg.geometry("700x520")
        self.dlg.resizable(False, False)
        self.dlg.transient(parent)
        self.dlg.grab_set()

        self.recibo_seleccionado = None
        self.var_combo_sel = tk.StringVar()
        self.var_fecha_pago = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.var_metodo_pago = tk.StringVar(value="Efectivo")
        self.var_monto_pago = tk.StringVar()
        self.var_saldar_total = tk.BooleanVar(value=True)
        self.var_obs = tk.StringVar()

        self._construir_ui(recibo_preseleccionado)
        self.var_metodo_pago.trace_add("write", self._al_cambiar_metodo_pago_dialogo)

    def _construir_ui(self, preseleccion):
        frame = ttk.Frame(self.dlg, padding="16 14 16 14")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame,
            text="Cobro de Saldo Pendiente",
            font=("Segoe UI", 12, "bold"),
            foreground="#0f172a",
        ).pack(anchor=tk.W)
        ttk.Label(
            frame,
            text="Seleccioná el recibo con deuda a cobrar. Al guardar, se actualizará el recibo original (quitando el color rojo/amarillo si queda saldado) y se asentará una línea en el Excel.",
            font=("Segoe UI", 9),
            foreground="#475569",
            wraplength=660,
        ).pack(anchor=tk.W, pady=(2, 10))

        # Selector de recibo
        f_sel = ttk.LabelFrame(frame, text=" 1. Recibo Pendiente a Imputar ", padding="10 8 10 8")
        f_sel.pack(fill=tk.X, pady=(0, 10))

        opciones_combo = []
        indice_a_elegir = 0
        for i, r in enumerate(self.recibos_con_deuda):
            txt = f"Recibo {r['nro_recibo']} - {r['nombre']} (Saldo: {formatear_moneda(r['saldo_pendiente'])}) - Emisión: {r['fecha_emision']}"
            opciones_combo.append(txt)
            if preseleccion is not None:
                if str(r.get("nro_recibo")).strip().lstrip("Rr") == str(preseleccion).strip().lstrip("Rr") or r.get("fila_excel") == preseleccion:
                    indice_a_elegir = i

        self.combo = ttk.Combobox(
            f_sel,
            textvariable=self.var_combo_sel,
            values=opciones_combo,
            state="readonly",
            font=("Segoe UI", 10),
        )
        self.combo.pack(fill=tk.X, pady=(2, 6))
        self.combo.bind("<<ComboboxSelected>>", self._al_cambiar_seleccion_combo)

        # Tarjeta informativa del recibo
        self.lbl_card_info = ttk.Label(
            f_sel,
            text="",
            font=("Segoe UI", 9),
            background="#F8FAFC",
            padding=6,
            relief="solid",
            borderwidth=1,
        )
        self.lbl_card_info.pack(fill=tk.X, pady=(2, 0))

        # Formulario de cobro
        f_cobro = ttk.LabelFrame(frame, text=" 2. Datos del Cobro Realizado ", padding="10 8 10 8")
        f_cobro.pack(fill=tk.X, pady=(0, 12))

        f_fila_c1 = ttk.Frame(f_cobro)
        f_fila_c1.pack(fill=tk.X, pady=(0, 8))

        f_fp = ttk.Frame(f_fila_c1)
        f_fp.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(f_fp, text="Fecha de este Cobro:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        f_fp_in = ttk.Frame(f_fp)
        f_fp_in.pack(anchor=tk.W, pady=(2, 0))
        self.entry_fecha_pago = ttk.Entry(
            f_fp_in, textvariable=self.var_fecha_pago, width=11, font=("Segoe UI", 10), justify="center"
        )
        self.entry_fecha_pago.pack(side=tk.LEFT, padx=(0, 2))
        self.btn_cal_fp = ttk.Button(
            f_fp_in,
            text="📅",
            width=3,
            command=lambda: abrir_calendario(self.btn_cal_fp, self.var_fecha_pago, permitir_vacio=False),
        )
        self.btn_cal_fp.pack(side=tk.LEFT, padx=(0, 4))
        self.entry_fecha_pago.bind(
            "<Double-Button-1>",
            lambda e: abrir_calendario(self.btn_cal_fp, self.var_fecha_pago, permitir_vacio=False),
        )
        ttk.Button(
            f_fp_in,
            text="Hoy",
            width=5,
            command=lambda: self.var_fecha_pago.set(date.today().strftime("%d/%m/%Y")),
        ).pack(side=tk.LEFT)

        f_mp = ttk.Frame(f_fila_c1)
        f_mp.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(f_mp, text="Método de Pago:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        f_mp_opts = ttk.Frame(f_mp)
        f_mp_opts.pack(anchor=tk.W, pady=(2, 0))
        for m in ("Efectivo", "Transferencia", "Débito"):
            ttk.Radiobutton(f_mp_opts, text=m, value=m, variable=self.var_metodo_pago).pack(side=tk.LEFT, padx=(0, 8))

        f_monto = ttk.Frame(f_fila_c1)
        f_monto.pack(side=tk.LEFT)
        ttk.Label(f_monto, text="Monto a Cobrar ($):", font=("Segoe UI", 9, "bold"), foreground="#0f172a").pack(anchor=tk.W)
        self.entry_monto = ttk.Entry(f_monto, textvariable=self.var_monto_pago, width=12, font=("Segoe UI", 11, "bold"), justify="right")
        self.entry_monto.pack(anchor=tk.W, pady=(2, 0))

        # Checkbox para cancelar deuda en su totalidad
        f_saldar = ttk.Frame(f_cobro)
        f_saldar.pack(fill=tk.X, pady=(2, 6))
        ttk.Checkbutton(
            f_saldar,
            text="Marcar deuda del recibo original como totalmente cancelada ($0 saldo pendiente)",
            variable=self.var_saldar_total,
        ).pack(side=tk.LEFT)

        # Observaciones
        f_obs = ttk.Frame(f_cobro)
        f_obs.pack(fill=tk.X)
        ttk.Label(f_obs, text="Observaciones adicionales (opcional):", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Entry(f_obs, textvariable=self.var_obs, font=("Segoe UI", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Botones inferiores
        f_btns = ttk.Frame(frame)
        f_btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(
            f_btns,
            text="💾 Registrar Cobro y Actualizar Excel",
            style="Save.TButton",
            command=self._guardar,
        ).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(f_btns, text="Cancelar", command=self.dlg.destroy).pack(side=tk.RIGHT)

        # Seleccionar por defecto
        self.combo.current(indice_a_elegir)
        self._al_cambiar_seleccion_combo()

    def _al_cambiar_seleccion_combo(self, event=None):
        idx = self.combo.current()
        if 0 <= idx < len(self.recibos_con_deuda):
            r = self.recibos_con_deuda[idx]
            self.recibo_seleccionado = r
            saldo = float(r["saldo_pendiente"])
            self.var_monto_pago.set(str(int(saldo) if saldo.is_integer() else saldo))

            txt = (
                f"Alumno: {r['nombre']}   |   Actividad: {r['actividad']}   |   Recibo N°: {r['nro_recibo']}\n"
                f"Total Original: {formatear_moneda(r['total'])}   |   Abonado Previo: {formatear_moneda(r['monto_pagado'])}   |   "
                f"DEUDA PENDIENTE: {formatear_moneda(saldo)} (Estado: {r['estado_pago']})"
            )
            self.lbl_card_info.config(text=txt)
            self._al_cambiar_metodo_pago_dialogo()

    def _al_cambiar_metodo_pago_dialogo(self, *args):
        if not self.recibo_seleccionado or not self.precios_data:
            return
        r = self.recibo_seleccionado
        mp_orig = str(r.get("metodo_pago", "")).strip().lower()

        # Si el método de pago original era 'No figura' o no estaba definido y debe todo el monto:
        pagado_previo = float(r.get("monto_pagado", 0.0) or 0.0)
        if mp_orig in ("no figura", "") and pagado_previo == 0:
            try:
                c_val = int(r.get("clases_sem") or 2)
            except (ValueError, TypeError):
                c_val = 2
            cuotas_dict = self.precios_data.get("cuotas", {}).get(c_val, {})
            desc_orig = str(r.get("descuento", "")).strip()
            nuevo_mp = self.var_metodo_pago.get()

            if desc_orig == "Jubilado":
                tipo_desc = "Jubilado"
            elif desc_orig == "Grupo familiar":
                tipo_desc = "Grupo familiar"
            elif nuevo_mp == "Efectivo":
                tipo_desc = "Efectivo"
            else:
                tipo_desc = "Ninguno"

            nueva_cuota = float(cuotas_dict.get(tipo_desc, cuotas_dict.get("Ninguno", 0.0)))
            insc_val = str(r.get("inscripcion", "")).strip()
            nueva_insc = float(self.precios_data.get("inscripcion", 15000.0)) if insc_val == "Sí" else 0.0
            sugerido = nueva_cuota + nueva_insc
            if sugerido > 0:
                self.var_monto_pago.set(str(int(sugerido) if sugerido.is_integer() else sugerido))

    def _guardar(self):
        if not self.recibo_seleccionado:
            messagebox.showwarning("Seleccionar Recibo", "Por favor seleccioná un recibo de la lista.", parent=self.dlg)
            return

        r = self.recibo_seleccionado
        dt_pago = parsear_fecha(self.var_fecha_pago.get())
        if not dt_pago:
            messagebox.showwarning("Fecha inválida", "Ingresá una fecha válida con formato DD/MM/AAAA.", parent=self.dlg)
            return

        try:
            limpio_monto = self.var_monto_pago.get().replace("$", "").replace(".", "").replace(",", ".").strip()
            monto_val = float(limpio_monto)
        except ValueError:
            messagebox.showwarning("Monto inválido", "El monto a cobrar ingresado no es válido.", parent=self.dlg)
            return

        if monto_val <= 0:
            messagebox.showwarning("Monto inválido", "El monto a cobrar debe ser mayor a 0.", parent=self.dlg)
            return

        saldo_max = float(r["saldo_pendiente"])
        if monto_val > saldo_max:
            conf = messagebox.askyesno(
                "Monto mayor a la deuda",
                f"El monto ingresado ({formatear_moneda(monto_val)}) es mayor a la deuda pendiente ({formatear_moneda(saldo_max)}).\n\n"
                f"¿Deseás registrar {formatear_moneda(saldo_max)} para saldar exactamente la deuda?",
                parent=self.dlg
            )
            if not conf:
                return
            monto_val = saldo_max

        try:
            self._registrar_fn(
                ruta_recibos=self.ruta_recibos,
                nro_recibo=r["nro_recibo"],
                nombre_alumno=r["nombre"],
                fecha_pago=dt_pago.strftime("%d/%m/%Y"),
                monto_pago=monto_val,
                metodo_pago=self.var_metodo_pago.get(),
                observaciones_extra=self.var_obs.get().strip(),
                fila_excel_especifica=r.get("fila_excel"),
                saldar_total=self.var_saldar_total.get(),
            )
        except PermissionError:
            messagebox.showerror(
                "Archivo Abierto",
                "No se pudo guardar porque 'recibos.xlsx' está abierto en Excel.\n\nPor favor cerralo y volvé a intentar.",
                parent=self.dlg
            )
            return
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"Ocurrió un error al registrar el pago:\n{e}", parent=self.dlg)
            return

        self.dlg.destroy()
        if self.on_guardar_callback:
            self.on_guardar_callback()

        messagebox.showinfo(
            "Pago Registrado",
            f"Se registró el cobro de {formatear_moneda(monto_val)} para {r['nombre']} exitosamente.\n\n"
            f"El recibo original fue actualizado y se asentó la línea correspondiente en el Excel.",
            parent=self.parent
        )


# ==========================================================
# DIÁLOGO MODAL PARA EMISIÓN DE CARNETS SIN RECIBO DE CUOTA
# ==========================================================
class DialogoCarnetSinRecibo:
    def __init__(
        self,
        parent,
        ruta_recibos,
        lista_alumnos,
        precio_carnet,
        proximo_carnet=1,
        on_guardar_callback=None,
        proximo_carnet_sugerido=None,
    ):
        self.parent = parent
        self.ruta_recibos = ruta_recibos
        self.lista_alumnos = lista_alumnos
        self.precio_carnet = float(precio_carnet or 1000.0)
        self.on_guardar_callback = on_guardar_callback

        if proximo_carnet_sugerido is not None:
            proximo_carnet = proximo_carnet_sugerido

        from excel_manager import registrar_carnet_sin_recibo
        self._registrar_fn = registrar_carnet_sin_recibo

        self.dlg = tk.Toplevel(parent)
        self.dlg.title("Emitir Carnet (Sin Recibo / Reposición)")
        self.dlg.geometry("640x450")
        self.dlg.resizable(False, False)
        self.dlg.transient(parent)
        self.dlg.grab_set()

        self.var_nombre = tk.StringVar()
        self.var_nro_carnet = tk.StringVar(value=str(proximo_carnet))
        self.var_recibo_asociado = tk.StringVar()
        self.var_fecha_recibo_asociado = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.var_fecha = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.var_modo_cobro = tk.StringVar(value="Cobrado")  # "Cobrado", "Sin cargo", "Debe"
        self.var_metodo_pago = tk.StringVar(value="Efectivo")
        self.var_obs = tk.StringVar()

        self._construir_ui()

    def _construir_ui(self):
        frame = ttk.Frame(self.dlg, padding="16 14 16 14")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame,
            text="Emisión de Carnet (Sin Recibo)",
            font=("Segoe UI", 12, "bold"),
            foreground="#0f172a",
        ).pack(anchor=tk.W)
        ttk.Label(
            frame,
            text="Registrá la emisión o reposición de un carnet independiente. Podés indicar si fue cobrado, sin cargo o adeudado.",
            font=("Segoe UI", 9),
            foreground="#475569",
        ).pack(anchor=tk.W, pady=(2, 10))

        f_form = ttk.LabelFrame(frame, text=" Datos del Carnet ", padding="10 8 10 8")
        f_form.pack(fill=tk.X, pady=(0, 12))

        # Fila 1: Nombre con autocomplete
        f_nom = ttk.Frame(f_form)
        f_nom.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(f_nom, text="Alumno (escribí para autocompletar):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        self.autocomplete = AutocompleteEntry(f_nom, textvariable=self.var_nombre)
        self.autocomplete.actualizar_base(self.lista_alumnos)
        self.autocomplete.pack(fill=tk.X, pady=(2, 0))

        # Fila 2: N° Carnet | Recibo Asociado | Fecha Recibo Asoc. | Fecha Emisión
        f_fila2 = ttk.Frame(f_form)
        f_fila2.pack(fill=tk.X, pady=(0, 8))

        f_car = ttk.Frame(f_fila2)
        f_car.pack(side=tk.LEFT, padx=(0, 14))
        ttk.Label(f_car, text="Carnet N°:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        ttk.Entry(f_car, textvariable=self.var_nro_carnet, width=8, font=("Segoe UI", 10, "bold"), justify="center").pack(anchor=tk.W, pady=(2, 0))

        f_rec_asoc = ttk.Frame(f_fila2)
        f_rec_asoc.pack(side=tk.LEFT, padx=(0, 14))
        ttk.Label(f_rec_asoc, text="Recibo Asoc. (ej. 2):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        ttk.Entry(f_rec_asoc, textvariable=self.var_recibo_asociado, width=15, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=(2, 0))

        f_fec_rec = ttk.Frame(f_fila2)
        f_fec_rec.pack(side=tk.LEFT, padx=(0, 14))
        ttk.Label(f_fec_rec, text="Fecha Recibo Asoc.:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        f_fec_rec_in = ttk.Frame(f_fec_rec)
        f_fec_rec_in.pack(anchor=tk.W, pady=(2, 0))
        self.entry_fecha_recibo = ttk.Entry(
            f_fec_rec_in, textvariable=self.var_fecha_recibo_asociado, width=11, font=("Segoe UI", 9), justify="center"
        )
        self.entry_fecha_recibo.pack(side=tk.LEFT, padx=(0, 2))
        self.btn_cal_fec_rec = ttk.Button(
            f_fec_rec_in,
            text="📅",
            width=3,
            command=lambda: abrir_calendario(self.btn_cal_fec_rec, self.var_fecha_recibo_asociado, permitir_vacio=True),
        )
        self.btn_cal_fec_rec.pack(side=tk.LEFT)
        self.entry_fecha_recibo.bind(
            "<Double-Button-1>",
            lambda e: abrir_calendario(self.btn_cal_fec_rec, self.var_fecha_recibo_asociado, permitir_vacio=True),
        )

        f_fec = ttk.Frame(f_fila2)
        f_fec.pack(side=tk.RIGHT)
        ttk.Label(f_fec, text="Fecha Emisión Carnet:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        f_fec_in = ttk.Frame(f_fec)
        f_fec_in.pack(anchor=tk.W, pady=(2, 0))
        self.entry_fecha = ttk.Entry(
            f_fec_in, textvariable=self.var_fecha, width=11, font=("Segoe UI", 9), justify="center"
        )
        self.entry_fecha.pack(side=tk.LEFT, padx=(0, 2))
        self.btn_cal_fec = ttk.Button(
            f_fec_in,
            text="📅",
            width=3,
            command=lambda: abrir_calendario(self.btn_cal_fec, self.var_fecha, permitir_vacio=False),
        )
        self.btn_cal_fec.pack(side=tk.LEFT, padx=(0, 4))
        self.entry_fecha.bind(
            "<Double-Button-1>",
            lambda e: abrir_calendario(self.btn_cal_fec, self.var_fecha, permitir_vacio=False),
        )
        ttk.Button(
            f_fec_in,
            text="Hoy",
            width=5,
            command=lambda: self.var_fecha.set(date.today().strftime("%d/%m/%Y")),
        ).pack(side=tk.LEFT)

        # Fila 3: Modalidad de cobro
        f_cobro_box = ttk.LabelFrame(f_form, text=" Condición de Cobro ", padding="8 6 8 6")
        f_cobro_box.pack(fill=tk.X, pady=(2, 8))

        precio_str = formatear_moneda(self.precio_carnet)

        f_opts_c = ttk.Frame(f_cobro_box)
        f_opts_c.pack(fill=tk.X)

        ttk.Radiobutton(
            f_opts_c,
            text=f"Cobrado ({precio_str})",
            value="Cobrado",
            variable=self.var_modo_cobro,
            command=self._al_cambiar_modo_cobro,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Radiobutton(
            f_opts_c,
            text="Sin cargo ($0 - No se cobra)",
            value="Sin cargo",
            variable=self.var_modo_cobro,
            command=self._al_cambiar_modo_cobro,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Radiobutton(
            f_opts_c,
            text=f"Debe ({precio_str} - Se marca en Rojo)",
            value="Debe",
            variable=self.var_modo_cobro,
            command=self._al_cambiar_modo_cobro,
        ).pack(side=tk.LEFT)

        # Método de pago (si es cobrado)
        self.f_metodo = ttk.Frame(f_cobro_box)
        self.f_metodo.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(self.f_metodo, text="Método de Pago:").pack(side=tk.LEFT, padx=(0, 8))
        for m in ("Efectivo", "Transferencia", "Débito"):
            ttk.Radiobutton(self.f_metodo, text=m, value=m, variable=self.var_metodo_pago).pack(side=tk.LEFT, padx=(0, 8))

        # Fila 4: Observaciones
        f_obs = ttk.Frame(f_form)
        f_obs.pack(fill=tk.X)
        ttk.Label(f_obs, text="Observaciones (opcional):", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Entry(f_obs, textvariable=self.var_obs, font=("Segoe UI", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Botones
        f_btns = ttk.Frame(frame)
        f_btns.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(
            f_btns,
            text="💾 Registrar Carnet en Excel",
            style="Save.TButton",
            command=self._guardar,
        ).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(f_btns, text="Cancelar", command=self.dlg.destroy).pack(side=tk.RIGHT)

        self.autocomplete.entry.focus_set()

    def _al_cambiar_modo_cobro(self):
        if self.var_modo_cobro.get() == "Cobrado":
            for child in self.f_metodo.winfo_children():
                try:
                    child.config(state="normal")
                except Exception:
                    pass
        else:
            for child in self.f_metodo.winfo_children():
                try:
                    child.config(state="disabled")
                except Exception:
                    pass

    def _guardar(self):
        nom = self.var_nombre.get().strip()
        if not nom:
            messagebox.showwarning("Falta Alumno", "Por favor ingresá el nombre del alumno.", parent=self.dlg)
            self.autocomplete.entry.focus_set()
            return

        car_str = self.var_nro_carnet.get().strip()
        if not car_str:
            messagebox.showwarning("Falta N° de Carnet", "Ingresá el número de carnet.", parent=self.dlg)
            return

        try:
            car_num = int(car_str)
        except ValueError:
            car_num = car_str

        dt_emision = parsear_fecha(self.var_fecha.get())
        if not dt_emision:
            messagebox.showwarning("Fecha inválida", "Ingresá una fecha válida DD/MM/AAAA.", parent=self.dlg)
            return

        rec_asoc = self.var_recibo_asociado.get().strip()
        fec_rec_asoc = self.var_fecha_recibo_asociado.get().strip()
        modo = self.var_modo_cobro.get()
        metodo = self.var_metodo_pago.get()
        obs = self.var_obs.get().strip()

        try:
            self._registrar_fn(
                ruta_recibos=self.ruta_recibos,
                nombre_alumno=nom,
                nro_carnet=car_num,
                recibo_asociado=rec_asoc,
                fecha_emision=dt_emision.strftime("%d/%m/%Y"),
                modo_cobro=modo,
                precio_carnet=self.precio_carnet,
                metodo_pago=metodo,
                observaciones_extra=obs,
                fecha_recibo_asociado=fec_rec_asoc,
            )
        except PermissionError:
            messagebox.showerror(
                "Archivo Abierto",
                "No se pudo guardar porque 'recibos.xlsx' está abierto en Excel.\n\nPor favor cerralo y volvé a intentar.",
                parent=self.dlg
            )
            return
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"Ocurrió un error al registrar el carnet:\n{e}", parent=self.dlg)
            return

        self.dlg.destroy()
        if self.on_guardar_callback:
            self.on_guardar_callback()

        messagebox.showinfo(
            "Carnet Registrado",
            f"Carnet N° {car_num} para {nom} registrado con éxito en el Excel.",
            parent=self.parent
        )
