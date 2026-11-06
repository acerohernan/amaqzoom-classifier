# -*- coding: utf-8 -*-
"""Amaqzoom · análisis de ventas y clientes."""
import math
import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from types import SimpleNamespace
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends._backend_tk import NavigationToolbar2Tk
import analisis as an
import graficos as gf

ROJO, AZUL, GRIS, FONDO = "#B50E2E", "#1F7A8C", "#555555", "#F4F4F4"
FUENTE = "Segoe UI" if sys.platform.startswith("win") else "Helvetica"


# --------------------------------------------------------------------------- #
# Widgets auxiliares
# --------------------------------------------------------------------------- #
class TablaDF(ttk.Frame):
    """Tabla (Treeview) que muestra un DataFrame, con orden por columna al hacer clic."""
    MAX_FILAS = 20000

    def __init__(self, master, altura=8):
        super().__init__(master)
        self.tree = ttk.Treeview(self, show="headings", height=altura)
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.df = pd.DataFrame()
        self._orden = (None, True)

    def mostrar(self, df: pd.DataFrame, indice=True):
        d = df.reset_index() if indice else df.copy()
        d.columns = [str(c) for c in d.columns]
        self.df = d
        self._orden = (None, True)
        self._pintar()

    def _ordenar(self, col):
        asc = not (self._orden[0] == col and self._orden[1])
        self.df = self.df.sort_values(col, ascending=asc, kind="stable")
        self._orden = (col, asc)
        self._pintar()

    @staticmethod
    def _fmt(v):
        if isinstance(v, (float, np.floating)):
            if pd.isna(v):
                return ""
            return f"{v:,.2f}" if abs(v) >= 10 else f"{v:.3f}"
        if isinstance(v, pd.Timestamp):
            return v.strftime("%d/%m/%Y")
        return "" if v is None else str(v)

    def _pintar(self):
        t = self.tree
        t.delete(*t.get_children())
        cols = list(self.df.columns)
        t["columns"] = cols
        for c in cols:
            marca = ""
            if self._orden[0] == c:
                marca = " ▲" if self._orden[1] else " ▼"
            t.heading(c, text=c + marca, command=lambda c=c: self._ordenar(c))
            muestra = self.df[c].head(200).map(self._fmt)
            largo = max([len(c) + 3] + [len(v) for v in muestra]) if len(muestra) else len(c)
            t.column(c, width=max(85, min(320, 8 * largo + 24)), anchor="center", stretch=True)
        for fila in self.df.head(self.MAX_FILAS).itertuples(index=False):
            t.insert("", "end", values=[self._fmt(v) for v in fila])


def kpi(master, valor, titulo, color=ROJO):
    f = tk.Frame(master, bg="white", highlightbackground="#DDDDDD", highlightthickness=1)
    tk.Label(f, text=str(valor), font=(FUENTE, 22, "bold"), fg=color, bg="white").pack(padx=18, pady=(10, 0))
    tk.Label(f, text=titulo, font=(FUENTE, 9), fg=GRIS, bg="white", wraplength=150).pack(padx=10, pady=(0, 10))
    return f


# --------------------------------------------------------------------------- #
# Aplicación
# --------------------------------------------------------------------------- #
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Amaqzoom · Análisis de ventas")
        self.geometry("1280x800")
        self.minsize(1050, 680)
        self.configure(bg=FONDO)
        self.report_callback_exception = self._excepcion_tk

        self.df_raw = None
        self.res = None
        self.origen = ""
        self.ruta = None
        self._cola = None
        self.knn_vars = {}
        self.knn_recencia = None
        self.knn_compras = None

        self._estilos()
        self._barra_superior()
        self._pestanas()
        self._barra_estado()
        self._mensaje_inicial()

    # ------------------------------------------------------------------ UI --
    def _estilos(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TFrame", background=FONDO)
        st.configure("TLabel", background=FONDO, font=(FUENTE, 10))
        st.configure("TCheckbutton", background=FONDO, font=(FUENTE, 10))
        st.configure("TButton", font=(FUENTE, 10), padding=6)
        st.configure("Accion.TButton", font=(FUENTE, 10, "bold"), foreground="white", background=ROJO)
        st.map("Accion.TButton", background=[("active", "#8F0B25"), ("disabled", "#C9A0A8")])
        st.configure("TNotebook.Tab", font=(FUENTE, 10), padding=(14, 7))
        st.configure("Treeview", font=(FUENTE, 9), rowheight=24)
        st.configure("Treeview.Heading", font=(FUENTE, 9, "bold"))
        st.configure("Titulo.TLabel", font=(FUENTE, 13, "bold"), foreground=ROJO)
        st.configure("Sub.TLabel", foreground=GRIS)

    def _barra_superior(self):
        top = ttk.Frame(self, padding=(12, 10, 12, 4))
        top.pack(fill="x")
        f1 = ttk.Frame(top)
        f1.pack(fill="x")
        ttk.Label(f1, text="Análisis de ventas", style="Titulo.TLabel").pack(side="left", padx=(0, 18))
        ttk.Button(f1, text="Importar ventas (Excel / CSV)…", command=self.cargar_archivo,
                   style="Accion.TButton").pack(side="left", padx=3)
        ttk.Label(f1, text="Hoja:").pack(side="left", padx=(10, 3))
        self.cb_hoja = ttk.Combobox(f1, state="disabled", width=18)
        self.cb_hoja.pack(side="left")
        self.cb_hoja.bind("<<ComboboxSelected>>", lambda e: self._leer_hoja())
        ttk.Button(f1, text="Datos de demostración", command=self.cargar_demo).pack(side="left", padx=(14, 3))
        ttk.Button(f1, text="Descargar plantilla", command=self.guardar_plantilla).pack(side="left", padx=3)
        ttk.Button(f1, text="Formato requerido", command=self.mostrar_formato).pack(side="left", padx=3)

        f2 = ttk.Frame(top)
        f2.pack(fill="x", pady=(8, 0))
        self.var_anon = tk.BooleanVar(value=True)
        ttk.Checkbutton(f2, text="Proteger identificadores al analizar", variable=self.var_anon).pack(side="left")
        self.btn_run = ttk.Button(f2, text="Analizar mis ventas", style="Accion.TButton",
                                  command=self.ejecutar, state="disabled")
        self.btn_run.pack(side="left", padx=(18, 3))
        self.btn_exp = ttk.Button(f2, text="Exportar resultados…", command=self.exportar, state="disabled")
        self.btn_exp.pack(side="left", padx=3)
        self.lbl_origen = ttk.Label(f2, text="Ningún dato cargado", style="Sub.TLabel")
        self.lbl_origen.pack(side="right")

    def _pestanas(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=12, pady=6)
        nombres = [("resumen", "Resumen del negocio"), ("perfil", "Tipos de clientes"),
                   ("clientes", "Detalle de clientes"), ("estac", "Temporadas de venta"),
                   ("datos", "Datos importados"), ("knn", "Clasificar nuevo cliente")]
        self.tab = {}
        for clave, texto in nombres:
            fr = ttk.Frame(self.nb, padding=8)
            self.nb.add(fr, text=texto)
            self.tab[clave] = fr

    def _barra_estado(self):
        bar = ttk.Frame(self, padding=(12, 0, 12, 8))
        bar.pack(fill="x")
        self.lbl_estado = ttk.Label(bar, text="Listo.", style="Sub.TLabel")
        self.lbl_estado.pack(side="left")
        self.progreso = ttk.Progressbar(bar, mode="indeterminate", length=180)
        self.progreso.pack(side="right")

    def _mensaje_inicial(self):
        for clave, fr in self.tab.items():
            self._vaciar(fr)
            ttk.Label(fr, style="Sub.TLabel", justify="center", font=(FUENTE, 11),
                      text=("Primero analice sus ventas para entrenar el modelo KNN."
                            if clave == "knn" else
                            "Importe su archivo de ventas para conocer cuánto vende,\n"
                            "quiénes son sus clientes y cuándo se concentran las ventas.\n\n"
                            "Necesita una fila por compra y las columnas fecha, cliente,\n"
                            "categoría y monto. También puede probar con datos de demostración.")
                      ).place(relx=.5, rely=.45, anchor="center")

    # ------------------------------------------------------------ utilidades --
    @staticmethod
    def _vaciar(frame):
        for w in frame.winfo_children():
            w.destroy()

    def _estado(self, texto):
        self.lbl_estado.config(text=texto)
        self.update_idletasks()

    def _excepcion_tk(self, exc, val, tb):
        messagebox.showerror("Error inesperado", "".join(traceback.format_exception(exc, val, tb))[-1500:])

    def _fig_en(self, parent, fig, **pack):
        cont = ttk.Frame(parent)
        cont.pack(**({"fill": "both", "expand": True} | pack))
        canvas = FigureCanvasTkAgg(fig, master=cont)
        barra = NavigationToolbar2Tk(canvas, cont, pack_toolbar=False)
        barra.pack(side="bottom", fill="x")
        widget = canvas.get_tk_widget()
        # matplotlib fija el tamaño del widget al de la figura (sobre todo con escalado de pantalla);
        # se pide un tamaño mínimo y se fuerza un redimensionado para que el gráfico encaje en el espacio.
        def ajustar(_e=None):
            if not widget.winfo_exists():
                return
            widget.configure(width=240, height=160)
            widget.after(40, reescalar)

        def reescalar():
            if widget.winfo_exists() and widget.winfo_width() > 10:
                canvas.resize(SimpleNamespace(width=widget.winfo_width(), height=widget.winfo_height()))

        widget.configure(width=240, height=160)
        widget.bind("<Map>", ajustar, add="+")
        widget.pack(side="top", fill="both", expand=True)
        canvas.draw_idle()
        return cont

    # ---------------------------------------------------------------- datos --
    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccione el archivo de ventas",
            filetypes=[("Excel o CSV", "*.xlsx *.xlsm *.xls *.csv"), ("Todos los archivos", "*.*")])
        if not ruta:
            return
        self.config(cursor="watch")
        self._estado("Leyendo archivo…")
        try:
            hojas = an.listar_hojas(ruta)
            df, hoja_ok, primer_error = None, None, None
            for h in hojas:                       # usa la primera hoja que tenga las columnas
                try:
                    df = an.leer_datos(ruta, h)
                    hoja_ok = h
                    break
                except an.ErrorDatos as e:
                    primer_error = primer_error or e
            if df is None:
                if primer_error is not None:
                    raise primer_error
                else:
                    raise an.ErrorDatos("No se encontró ninguna hoja con las columnas requeridas.")
        except an.ErrorDatos as e:
            messagebox.showerror("Archivo no válido", str(e))
            self._estado("Listo.")
            return
        except Exception as e:                    # archivo corrupto, abierto en Excel, etc.
            messagebox.showerror("No se pudo leer el archivo", f"{type(e).__name__}: {e}")
            self._estado("Listo.")
            return
        finally:
            self.config(cursor="")
        self.ruta = ruta
        self.cb_hoja.config(state="readonly" if len(hojas) > 1 else "disabled", values=hojas)
        self.cb_hoja.set(hoja_ok)
        self._fijar_datos(df, f"{os.path.basename(ruta)} · hoja «{hoja_ok}»")

    def _leer_hoja(self):
        try:
            df = an.leer_datos(self.ruta or '', self.cb_hoja.get())
        except an.ErrorDatos as e:
            messagebox.showerror("Hoja no válida", str(e))
            return
        self._fijar_datos(df, f"{os.path.basename(self.ruta or '')} · hoja «{self.cb_hoja.get()}»")

    def cargar_demo(self):
        self.config(cursor="watch")
        self._estado("Generando datos sintéticos…")
        df = an.generar_datos_sinteticos()
        self.config(cursor="")
        self.cb_hoja.config(state="disabled", values=[])
        self.cb_hoja.set("")
        self.ruta = None
        self._fijar_datos(df, "Datos sintéticos de demostración")

    def _fijar_datos(self, df, origen):
        self.df_raw, self.res = df, None
        self.lbl_origen.config(text=f"{origen} · {len(df):,} filas")
        self.btn_run.config(state="normal")
        self.btn_exp.config(state="disabled")
        for clave, fr in self.tab.items():
            if clave != "datos":
                self._vaciar(fr)
                ttk.Label(fr, style="Sub.TLabel", font=(FUENTE, 11),
                          text="Presione «Analizar mis ventas» para ver los resultados."
                          ).place(relx=.5, rely=.45, anchor="center")
        self._tab_clasificar(None)
        fr = self.tab["datos"]
        self._vaciar(fr)
        ttk.Label(fr, text=f"Vista previa de las primeras 500 filas ({len(df):,} en total)",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
        t = TablaDF(fr, altura=20)
        t.pack(fill="both", expand=True)
        t.mostrar(df.head(500), indice=False)
        self.nb.select(self.tab["datos"])
        self._estado("Archivo listo. Presione «Analizar mis ventas».")

    def guardar_plantilla(self):
        ruta = filedialog.asksaveasfilename(title="Guardar plantilla", defaultextension=".xlsx",
                                            initialfile="plantilla_ventas_amaqzoom.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            an.plantilla_ejemplo().to_excel(ruta, index=False, sheet_name="Ventas")
        except Exception as e:
            messagebox.showerror("No se pudo guardar", str(e))
            return
        messagebox.showinfo("Plantilla guardada", f"Se guardó la plantilla en:\n{ruta}")

    def mostrar_formato(self):
        messagebox.showinfo(
            "Formato requerido",
            "El Excel (o CSV) debe tener una fila por transacción, con estas columnas:\n\n"
            "  • fecha        → ej. 05/10/2025 (día/mes/año) o fecha de Excel\n"
            "  • cliente      → identificador del cliente (DNI, teléfono, código…)\n"
            "  • canal        → opcional (tienda, Sodimac, etc.)\n"
            "  • categoria    → ej. Ventiladores, Cocina, Audio\n"
            "  • monto        → número o texto tipo «S/ 1,250.50»\n\n"
            "Los nombres no distinguen mayúsculas ni tildes.\n"
            "También se aceptan nombres como id_cliente, importe, total, línea o familia.\n\n"
            "La limpieza (duplicados, nulos, montos con «S/», categorías en mayúsculas) "
            "se hace automáticamente.\n\nSi hay más de 6 categorías, las de menor gasto se agrupan en «Otros».")

    # -------------------------------------------------------------- análisis --
    def ejecutar(self):
        if self.df_raw is None:
            return
        anon, df = self.var_anon.get(), self.df_raw.copy()
        self._cola = q = queue.Queue()
        self.btn_run.config(state="disabled")
        self.btn_exp.config(state="disabled")
        self.progreso.start(12)
        self.config(cursor="watch")

        def trabajo():
            try:
                res = an.ejecutar_analisis(df, anon, progreso=lambda m: q.put(("msg", m)))
                q.put(("ok", res))
            except Exception as e:
                q.put(("err", e, traceback.format_exc()))

        threading.Thread(target=trabajo, daemon=True).start()
        self.after(120, self._sondear)

    def _sondear(self):
        try:
            while True:
                if self._cola is None:
                    return
                m = self._cola.get_nowait()
                if m[0] == "msg":
                    self._estado(m[1])
                elif m[0] == "ok":
                    self._terminar()
                    self.res = m[1]
                    self._mostrar_resultados()
                    return
                else:
                    self._terminar()
                    if isinstance(m[1], an.ErrorDatos):
                        messagebox.showerror("No se pudo completar el análisis", str(m[1]))
                    else:
                        messagebox.showerror("Error en el análisis", m[2][-1500:])
                    self._estado("Error en el análisis.")
                    return
        except queue.Empty:
            pass
        self.after(120, self._sondear)

    def _terminar(self):
        self.progreso.stop()
        self.config(cursor="")
        self.btn_run.config(state="normal")

    # ------------------------------------------------------------ resultados --
    def _mostrar_resultados(self):
        r = self.res
        if r is None:
            return
        self._tab_resumen(r)
        self._tab_perfil(r)
        self._tab_clientes(r)
        self._tab_estacionalidad(r)
        self._tab_clasificar(r)
        self.btn_exp.config(state="normal")
        self.nb.select(self.tab["resumen"])
        self._estado(f"Listo: se analizaron {len(r.ventas):,} compras y {r.resumen['clientes']:,} clientes.")

    def _tab_resumen(self, r):
        fr = self.tab["resumen"]
        self._vaciar(fr)
        fila = ttk.Frame(fr)
        fila.pack(fill="x", pady=(0, 10))
        ventas_totales = float(r.ventas["monto"].sum())
        n_compras = len(r.ventas)
        ticket_promedio = ventas_totales / n_compras if n_compras else 0
        clientes_recurrentes = int((r.clientes["frecuencia"] > 1).sum())
        porcentaje_recurrentes = clientes_recurrentes / len(r.clientes) if len(r.clientes) else 0
        tarjetas = [
            (f"S/ {ventas_totales:,.2f}", "ventas registradas", ROJO),
            (f"{r.resumen['clientes']:,}", "clientes atendidos", AZUL),
            (f"{n_compras:,}", "compras válidas", "#E8A33D"),
            (f"S/ {ticket_promedio:,.2f}", "venta promedio por compra", ROJO),
            (f"{porcentaje_recurrentes:.0%}", "clientes con más de una compra", AZUL),
        ]
        for i, (v, t, c) in enumerate(tarjetas):
            kpi(fila, v, t, c).grid(row=0, column=i, padx=5, sticky="nsew")
            fila.columnconfigure(i, weight=1, uniform="k")

        cuerpo = ttk.Frame(fr)
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=3)
        cuerpo.columnconfigure(1, weight=2)
        cuerpo.rowconfigure(0, weight=1)

        der = ttk.Frame(cuerpo)
        der.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        self._fig_en(der, gf.fig_tamano_segmentos(r))

        izq = ttk.Frame(cuerpo)
        izq.grid(row=0, column=0, sticky="nsew")
        ttk.Label(izq, text="Lo más importante de tus ventas", style="Titulo.TLabel").pack(anchor="w", pady=(0, 8))
        periodo = (f"Periodo analizado: {r.ventas['fecha'].min():%d/%m/%Y} – "
                   f"{r.ventas['fecha'].max():%d/%m/%Y}")
        filas_descartadas = r.informe["registros_iniciales"] - r.informe["registros_finales"]
        nota_limpieza = (f"Se revisaron {r.informe['registros_iniciales']:,} filas y se utilizaron "
                         f"{r.informe['registros_finales']:,} compras válidas. "
                         f"{filas_descartadas:,} filas incompletas, repetidas o inválidas no se incluyeron.")
        ttk.Label(izq, text=periodo, style="Sub.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(izq, text=nota_limpieza, style="Sub.TLabel", wraplength=650,
                  justify="left").pack(anchor="w", pady=(0, 14))

        por_ventas = r.perfil.assign(ventas_estimadas=r.perfil["clientes"] * r.perfil["monto_medio"])
        segmento_valor = por_ventas.loc[por_ventas["ventas_estimadas"].idxmax()]
        segmento_inactivo = r.perfil.loc[r.perfil["recencia_media"].idxmax()]
        categoria_ventas = r.ventas.groupby("categoria")["monto"].sum().sort_values(ascending=False)
        categoria_top = categoria_ventas.index[0]
        participacion = categoria_ventas.iloc[0] / ventas_totales if ventas_totales else 0
        recomendaciones = [
            f"Mayor aporte observado: {gf.nombre_cliente(segmento_valor['nombre'])}. "
            f"Representa {segmento_valor['%_clientes']:.1f}% de tus clientes y genera "
            f"aprox. S/ {segmento_valor['ventas_estimadas']:,.2f} en ventas registradas.",
            f"Oportunidad de retorno: {gf.nombre_cliente(segmento_inactivo['nombre'])} "
            f"lleva en promedio {segmento_inactivo['recencia_media']:.0f} días sin comprar. "
            "Considera una campaña de reactivación.",
            f"Categoría con más ventas: {categoria_top} ({participacion:.0%} del total). "
            "Revisa su disponibilidad y planifica promociones para mantener ese resultado.",
        ]
        for texto in recomendaciones:
            ttk.Label(izq, text="• " + texto, style="Sub.TLabel", wraplength=650,
                      justify="left").pack(anchor="w", pady=(0, 10))
        privacidad = ("Los identificadores se protegieron antes de analizar."
                      if r.anonimizado else
                      "Los identificadores originales se conservaron durante el análisis.")
        ttk.Label(izq, text=privacidad, style="Sub.TLabel", wraplength=650,
                  justify="left").pack(anchor="w", pady=(6, 0))

    def _tab_perfil(self, r):
        fr = self.tab["perfil"]
        self._vaciar(fr)
        ttk.Label(fr, text="Conoce los distintos tipos de clientes y qué compran",
                  style="Titulo.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(fr, text="Las ventas por cliente corresponden al periodo incluido en el archivo.",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
        self._fig_en(fr, gf.fig_perfil(r))
        tabla = r.perfil[["nombre", "clientes", "%_clientes", "recencia_media",
                          "frecuencia_media", "monto_medio"] + r.cols_pct].copy()
        tabla = tabla.rename(columns={
            "nombre": "Tipo de cliente",
            "clientes": "N.º de clientes",
            "%_clientes": "% de clientes",
            "recencia_media": "Días desde última compra",
            "frecuencia_media": "Compras por cliente",
            "monto_medio": "Ventas por cliente (S/)",
            **{c: f"Gasto en {categoria} (%)" for c, categoria in zip(r.cols_pct, r.categorias)},
        })
        tabla["Tipo de cliente"] = tabla["Tipo de cliente"].map(gf.nombre_cliente)
        tabla["% de clientes"] = tabla["% de clientes"].map(lambda v: f"{v:.1f}%")
        for categoria in r.categorias:
            col = f"Gasto en {categoria} (%)"
            tabla[col] = tabla[col].map(lambda v: f"{v:.0%}")
        tabla["Ventas por cliente (S/)"] = tabla["Ventas por cliente (S/)"].map(lambda v: f"S/ {v:,.2f}")
        tabla["Compras por cliente"] = tabla["Compras por cliente"].map(lambda v: f"{v:.1f}")
        tabla["Días desde última compra"] = tabla["Días desde última compra"].map(lambda v: f"{v:.0f}")
        t = TablaDF(fr, altura=min(8, len(tabla) + 1))
        t.pack(side="bottom", fill="x", pady=(6, 0))
        t.mostrar(tabla, indice=False)

    def _tab_clientes(self, r):
        fr = self.tab["clientes"]
        self._vaciar(fr)
        ttk.Label(fr, text="Consulta las compras y el perfil de cada cliente",
                  style="Titulo.TLabel").pack(anchor="w", pady=(0, 2))
        ttk.Label(fr, text="Los importes reflejan las ventas registradas en el periodo analizado.",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 6))
        barra = ttk.Frame(fr)
        barra.pack(fill="x", pady=(0, 6))
        ttk.Label(barra, text="Tipo de cliente:").pack(side="left")
        nombres = sorted(set(r.nombres_seg.values()))
        opciones = ["Todos"] + [gf.nombre_cliente(nombre) for nombre in nombres]
        cb = ttk.Combobox(barra, state="readonly", values=opciones, width=38)
        cb.current(0)
        cb.pack(side="left", padx=6)
        lbl = ttk.Label(barra, style="Sub.TLabel")
        lbl.pack(side="left", padx=10)
        tabla = TablaDF(fr, altura=20)
        tabla.pack(fill="both", expand=True)

        clientes = r.clientes.reset_index().copy()
        clientes["nombre_segmento"] = clientes["nombre_segmento"].map(gf.nombre_cliente)
        columnas = ["id_cliente", "nombre_segmento", "recencia", "frecuencia", "monto"] + r.cols_pct
        nombres_columnas = {
            "id_cliente": "ID de cliente",
            "nombre_segmento": "Tipo de cliente",
            "recencia": "Días desde última compra",
            "frecuencia": "Compras registradas",
            "monto": "Ventas registradas (S/)",
            **{c: f"Gasto en {categoria} (%)" for c, categoria in zip(r.cols_pct, r.categorias)},
        }

        def filtrar(_=None):
            nombre = cb.get()
            datos = clientes if nombre == "Todos" else clientes[clientes["nombre_segmento"] == nombre]
            lbl.config(text=f"{len(datos):,} clientes · ordenados por ventas")
            datos = datos.sort_values("monto", ascending=False)[columnas].rename(columns=nombres_columnas)
            datos["Ventas registradas (S/)"] = datos["Ventas registradas (S/)"].map(
                lambda v: f"S/ {v:,.2f}")
            for categoria in r.categorias:
                col = f"Gasto en {categoria} (%)"
                datos[col] = datos[col].map(lambda v: f"{v:.0%}")
            tabla.mostrar(datos, indice=False)

        cb.bind("<<ComboboxSelected>>", filtrar)
        filtrar()

    def _tab_estacionalidad(self, r):
        fr = self.tab["estac"]
        self._vaciar(fr)
        ttk.Label(fr, text="Identifica cuándo preparar tus productos y promociones",
                  style="Titulo.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(fr, text="100% representa las ventas mensuales promedio de cada categoría.",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
        t = TablaDF(fr, altura=min(8, len(r.calendario) + 1))
        t.pack(side="bottom", fill="x", pady=(6, 0))
        calendario = r.calendario.rename(columns={
            "categoria": "Categoría",
            "meses_pico": "Meses con más ventas",
            "indice_pico": "Ventas vs. promedio",
            "accion": "Sugerencia",
        }).copy()
        calendario["Ventas vs. promedio"] = calendario["Ventas vs. promedio"].map(
            lambda v: f"{v:.0%} del promedio mensual")
        calendario["Sugerencia"] = calendario["Sugerencia"].str.replace(
            "Campaña", "Prepara una campaña", regex=False).str.replace(
                " y refuerzo de stock", " y refuerza el inventario", regex=False)
        t.mostrar(calendario, indice=False)
        self._fig_en(fr, gf.fig_estacionalidad(r))

    def _tab_clasificar(self, r):
        fr = self.tab["knn"]
        self._vaciar(fr)
        if r is None:
            ttk.Label(fr, text="Analice primero sus ventas para preparar el modelo KNN.",
                      style="Sub.TLabel", font=(FUENTE, 11)).place(
                          relx=.5, rely=.45, anchor="center")
            return

        ttk.Label(fr, text="Conoce el perfil probable de un cliente nuevo",
                  style="Titulo.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(
            fr, text=("Indique los días desde su última compra, sus compras registradas y "
                      "cuánto ha gastado por categoría. Los importes deben ser mayores "
                      "o iguales a cero y sumar más de cero."),
            style="Sub.TLabel", wraplength=900, justify="left").pack(anchor="w", pady=(0, 12))

        contenido = ttk.Frame(fr)
        contenido.pack(fill="both", expand=True)
        formulario = ttk.LabelFrame(contenido, text="Datos del cliente", padding=14)
        formulario.pack(side="left", fill="y", padx=(0, 14), anchor="n")

        self.knn_recencia = tk.StringVar()
        self.knn_compras = tk.StringVar()
        self.knn_vars = {}

        ttk.Label(formulario, text="Días desde la última compra:").grid(
            row=0, column=0, sticky="w", pady=5)
        ttk.Entry(formulario, textvariable=self.knn_recencia, width=18).grid(
            row=0, column=1, sticky="ew", padx=(12, 0), pady=5)
        ttk.Label(formulario, text="Número de compras:").grid(
            row=1, column=0, sticky="w", pady=5)
        ttk.Entry(formulario, textvariable=self.knn_compras, width=18).grid(
            row=1, column=1, sticky="ew", padx=(12, 0), pady=5)

        ttk.Separator(formulario).grid(row=2, column=0, columnspan=2, sticky="ew", pady=8)
        ttk.Label(formulario, text="Gasto registrado por categoría (S/):",
                  style="Titulo.TLabel").grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 5))
        for i, categoria in enumerate(r.categorias, start=4):
            var = tk.StringVar(value="0")
            self.knn_vars[categoria] = var
            ttk.Label(formulario, text=f"{categoria}:").grid(row=i, column=0, sticky="w", pady=3)
            ttk.Entry(formulario, textvariable=var, width=18).grid(
                row=i, column=1, sticky="ew", padx=(12, 0), pady=3)

        ttk.Button(formulario, text="Clasificar cliente", style="Accion.TButton",
                   command=self.clasificar_cliente).grid(
                       row=4 + len(r.categorias), column=0, columnspan=2,
                       sticky="ew", pady=(12, 0))
        formulario.columnconfigure(1, weight=1)

        self.knn_resultado = ttk.LabelFrame(contenido, text="Resultado", padding=18)
        self.knn_resultado.pack(side="left", fill="both", expand=True, anchor="n")
        self.knn_resultado_texto = ttk.Label(
            self.knn_resultado,
            text="Complete los datos y seleccione «Clasificar cliente».",
            style="Sub.TLabel", justify="left", wraplength=500)
        self.knn_resultado_texto.pack(anchor="nw")

    def clasificar_cliente(self):
        if self.res is None or self.knn_recencia is None or self.knn_compras is None:
            return
        try:
            recencia = int(self.knn_recencia.get())
            compras = int(self.knn_compras.get())
            gastos = {categoria: float(var.get()) for categoria, var in self.knn_vars.items()}
        except ValueError:
            messagebox.showerror(
                "Datos no válidos",
                "Ingrese días y compras como números enteros, y los gastos como importes numéricos.")
            return

        if recencia < 0:
            messagebox.showerror("Datos no válidos", "Los días desde la última compra no pueden ser negativos.")
            return
        if compras < 1:
            messagebox.showerror("Datos no válidos", "El número de compras debe ser al menos 1.")
            return
        if any(not math.isfinite(gasto) or gasto < 0 for gasto in gastos.values()):
            messagebox.showerror("Datos no válidos", "Los gastos deben ser importes válidos mayores o iguales a cero.")
            return

        try:
            _, segmento, confianza = self.res.clasificar_cliente(recencia, compras, gastos)
        except an.ErrorDatos as e:
            messagebox.showerror("No se pudo clasificar", str(e))
            return

        self.knn_resultado_texto.config(
            text=(f"Perfil probable: {gf.nombre_cliente(segmento)}\n\n"
                  f"Confianza estimada por el modelo: {confianza:.0%}\n\n"
                  "La clasificación es una estimación basada en los clientes del periodo analizado."))
        self._estado(f"Cliente clasificado: {gf.nombre_cliente(segmento)}.")

    # ------------------------------------------------------------- exportar --
    def exportar(self):
        if self.res is None:
            return
        base = filedialog.askdirectory(title="Seleccione dónde guardar los resultados")
        if not base:
            return
        carpeta = os.path.join(base, "resultados_amaqzoom_" + datetime.now().strftime("%Y%m%d_%H%M"))
        self.config(cursor="watch")
        self._estado("Exportando resultados…")
        try:
            an.exportar_tablas(self.res, carpeta)
            gf.guardar_figuras(self.res, carpeta)
        except PermissionError:
            messagebox.showerror("No se pudo exportar", "Sin permisos de escritura, o un archivo está abierto en Excel.")
            return
        except Exception as e:
            messagebox.showerror("No se pudo exportar", f"{type(e).__name__}: {e}")
            return
        finally:
            self.config(cursor="")
            self._estado("Listo.")
        if messagebox.askyesno("Exportación completa",
                               f"Resultados guardados en:\n{carpeta}\n\n¿Abrir la carpeta?"):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(carpeta)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", carpeta])
                else:
                    subprocess.Popen(["xdg-open", carpeta])
            except Exception:
                pass


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
