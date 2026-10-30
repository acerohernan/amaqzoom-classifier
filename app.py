# -*- coding: utf-8 -*-
"""
Amaqzoom · Segmentación de clientes (RFM + K-Means + KNN)
Aplicación de escritorio - ejecutar con:  python app.py
"""
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
        self.title("Amaqzoom · Segmentación de clientes (RFM + K-Means + KNN)")
        self.geometry("1280x800")
        self.minsize(1050, 680)
        self.configure(bg=FONDO)
        self.report_callback_exception = self._excepcion_tk

        self.df_raw = None
        self.res = None
        self.origen = ""
        self.ruta = None
        self._cola = None

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
        ttk.Label(f1, text="Segmentación de clientes", style="Titulo.TLabel").pack(side="left", padx=(0, 18))
        ttk.Button(f1, text="Cargar Excel / CSV…", command=self.cargar_archivo).pack(side="left", padx=3)
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
        ttk.Checkbutton(f2, text="Anonimizar id_cliente (SHA-256)", variable=self.var_anon).pack(side="left")
        ttk.Label(f2, text="Número de grupos (k):").pack(side="left", padx=(18, 3))
        self.cb_k = ttk.Combobox(f2, state="readonly", width=11,
                                 values=["Automático"] + [str(k) for k in range(an.K_MIN, an.K_MAX + 1)])
        self.cb_k.set("Automático")
        self.cb_k.pack(side="left")
        self.btn_run = ttk.Button(f2, text="▶  Ejecutar análisis", style="Accion.TButton",
                                  command=self.ejecutar, state="disabled")
        self.btn_run.pack(side="left", padx=(18, 3))
        self.btn_exp = ttk.Button(f2, text="Exportar resultados…", command=self.exportar, state="disabled")
        self.btn_exp.pack(side="left", padx=3)
        self.lbl_origen = ttk.Label(f2, text="Ningún dato cargado", style="Sub.TLabel")
        self.lbl_origen.pack(side="right")

    def _pestanas(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=12, pady=6)
        nombres = [("datos", "Datos cargados"), ("resumen", "Resumen"), ("k", "Selección de k"),
                   ("pca", "Segmentos (PCA)"), ("perfil", "Perfil RFM"), ("clientes", "Clientes"),
                   ("knn", "Clasificador KNN"), ("estac", "Estacionalidad"), ("nuevo", "Cliente nuevo")]
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
                      text=("Cargue un Excel con las columnas  fecha · id_cliente · canal · categoria · monto\n"
                            "y presione «Ejecutar análisis».\n\n"
                            "Si solo quiere probar, use «Datos de demostración».")
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
                          text="Presione «Ejecutar análisis» para ver los resultados."
                          ).place(relx=.5, rely=.45, anchor="center")
        fr = self.tab["datos"]
        self._vaciar(fr)
        ttk.Label(fr, text=f"Vista previa de las primeras 500 filas ({len(df):,} en total)",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
        t = TablaDF(fr, altura=20)
        t.pack(fill="both", expand=True)
        t.mostrar(df.head(500), indice=False)
        self.nb.select(self.tab["datos"])
        self._estado("Datos cargados. Presione «Ejecutar análisis».")

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
            "  • id_cliente   → identificador del cliente (DNI, teléfono, código…)\n"
            "  • canal        → opcional (tienda, Sodimac, etc.)\n"
            "  • categoria    → ej. Ventiladores, Cocina, Audio\n"
            "  • monto        → número o texto tipo «S/ 1,250.50»\n\n"
            "Los nombres no distinguen mayúsculas ni tildes.\n"
            "También se aceptan: cliente, importe, total, línea, familia…\n\n"
            "La limpieza (duplicados, nulos, montos con «S/», categorías en mayúsculas) "
            "se hace automáticamente.\n\nSi hay más de 6 categorías, las de menor gasto se agrupan en «Otros».")

    # -------------------------------------------------------------- análisis --
    def ejecutar(self):
        if self.df_raw is None:
            return
        k_txt = self.cb_k.get()
        k_manual = None if k_txt == "Automático" else int(k_txt)
        anon, df = self.var_anon.get(), self.df_raw.copy()
        self._cola = q = queue.Queue()
        self.btn_run.config(state="disabled")
        self.btn_exp.config(state="disabled")
        self.progreso.start(12)
        self.config(cursor="watch")

        def trabajo():
            try:
                res = an.ejecutar_analisis(df, anon, k_manual, progreso=lambda m: q.put(("msg", m)))
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
        self._tab_k(r)
        self._tab_pca(r)
        self._tab_perfil(r)
        self._tab_clientes(r)
        self._tab_knn(r)
        self._tab_estacionalidad(r)
        self._tab_nuevo(r)
        self.btn_exp.config(state="normal")
        self.nb.select(self.tab["resumen"])
        self._estado(f"Análisis completado: {r.resumen['clientes']:,} clientes en {r.k_opt} segmentos.")

    def _tab_resumen(self, r):
        fr = self.tab["resumen"]
        self._vaciar(fr)
        fila = ttk.Frame(fr)
        fila.pack(fill="x", pady=(0, 10))
        tarjetas = [
            (f"{r.resumen['clientes']:,}", "clientes analizados", ROJO),
            (r.k_opt, "segmentos (k) " + ("— automático" if r.k_automatico else "— manual"), ROJO),
            (f"{r.silhouette:.3f}", "silhouette (más alto = grupos más separados)", AZUL),
            (f"{r.davies_bouldin:.3f}", "Davies-Bouldin (más bajo = mejor)", AZUL),
            (f"{r.exactitud:.1%}", f"exactitud KNN (K={r.k_knn})", "#E8A33D"),
            (f"{r.f1_macro:.3f}", "F1 macro del KNN", "#E8A33D"),
        ]
        for i, (v, t, c) in enumerate(tarjetas):
            kpi(fila, v, t, c).grid(row=0, column=i, padx=5, sticky="nsew")
            fila.columnconfigure(i, weight=1, uniform="k")

        cuerpo = ttk.Frame(fr)
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=2)
        cuerpo.columnconfigure(1, weight=3)
        cuerpo.rowconfigure(0, weight=1)

        izq = ttk.Frame(cuerpo)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ttk.Label(izq, text="Informe de limpieza (HU02)", style="Titulo.TLabel").pack(anchor="w")
        etiquetas = {"registros_iniciales": "Registros iniciales",
                     "duplicados_eliminados": "Duplicados eliminados",
                     "nulos_eliminados": "Filas con nulos eliminadas",
                     "montos_fechas_invalidos": "Montos / fechas inválidos",
                     "registros_finales": "Registros finales",
                     "clientes_unicos": "Clientes únicos"}
        inf = pd.DataFrame({"concepto": [etiquetas[k] for k in r.informe],
                            "valor": [f"{v:,}" for v in r.informe.values()]})
        t = TablaDF(izq, altura=7)
        t.pack(fill="x", pady=(2, 8))
        t.mostrar(inf, indice=False)
        nota = ("IDs anonimizados con SHA-256 + sal." if r.anonimizado
                else "IDs originales conservados (sin anonimizar).")
        ttk.Label(izq, text=nota, style="Sub.TLabel").pack(anchor="w")
        ttk.Label(izq, style="Sub.TLabel", wraplength=380, justify="left",
                  text=f"Categorías usadas: {', '.join(r.categorias)}.").pack(anchor="w", pady=(6, 0))

        der = ttk.Frame(cuerpo)
        der.grid(row=0, column=1, sticky="nsew")
        self._fig_en(der, gf.fig_tamano_segmentos(r))

    def _tab_k(self, r):
        fr = self.tab["k"]
        self._vaciar(fr)
        tx = ("elegido automáticamente por mayor silhouette" if r.k_automatico else "fijado manualmente")
        ttk.Label(fr, text=f"k = {r.k_opt}  ({tx})", style="Titulo.TLabel").pack(anchor="w")
        t = TablaDF(fr, altura=7)
        t.pack(side="bottom", fill="x", pady=(6, 0))
        t.mostrar(r.tabla_k.round(3), indice=False)
        self._fig_en(fr, gf.fig_codo_silhouette(r))

    def _tab_pca(self, r):
        fr = self.tab["pca"]
        self._vaciar(fr)
        self._fig_en(fr, gf.fig_pca(r))

    def _tab_perfil(self, r):
        fr = self.tab["perfil"]
        self._vaciar(fr)
        t = TablaDF(fr, altura=min(8, len(r.perfil) + 1))
        t.pack(side="bottom", fill="x", pady=(6, 0))
        t.mostrar(r.perfil[["nombre", "clientes", "%_clientes", "recencia_media", "frecuencia_media",
                            "monto_medio"] + r.cols_pct])
        self._fig_en(fr, gf.fig_perfil(r))

    def _tab_clientes(self, r):
        fr = self.tab["clientes"]
        self._vaciar(fr)
        barra = ttk.Frame(fr)
        barra.pack(fill="x", pady=(0, 6))
        ttk.Label(barra, text="Segmento:").pack(side="left")
        opciones = ["Todos"] + [f"S{s} · {r.nombres_seg[s]}" for s in sorted(r.nombres_seg)]
        cb = ttk.Combobox(barra, state="readonly", values=opciones, width=34)
        cb.set("Todos")
        cb.pack(side="left", padx=6)
        lbl = ttk.Label(barra, style="Sub.TLabel")
        lbl.pack(side="left", padx=10)
        tabla = TablaDF(fr, altura=20)
        tabla.pack(fill="both", expand=True)

        def filtrar(_=None):
            d = r.clientes
            if cb.current() > 0:
                d = d[d["segmento"] == sorted(r.nombres_seg)[cb.current() - 1]]
            lbl.config(text=f"{len(d):,} clientes (clic en una columna para ordenar)")
            tabla.mostrar(d.sort_values("monto", ascending=False))

        cb.bind("<<ComboboxSelected>>", filtrar)
        filtrar()

    def _tab_knn(self, r):
        fr = self.tab["knn"]
        self._vaciar(fr)
        cv = ", ".join(f"K={k}: {v:.3f}" for k, v in r.cv_knn.items()) or "no aplicable (muestra pequeña)"
        ttk.Label(fr, style="Titulo.TLabel",
                  text=f"KNN con K={r.k_knn} · exactitud {r.exactitud:.1%} · F1 macro {r.f1_macro:.3f}"
                  ).pack(anchor="w")
        ttk.Label(fr, style="Sub.TLabel", wraplength=1100, justify="left",
                  text=f"Validación cruzada (exactitud): {cv}").pack(anchor="w", pady=(0, 4))
        cuerpo = ttk.Frame(fr)
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=1)
        cuerpo.columnconfigure(1, weight=1)
        cuerpo.rowconfigure(0, weight=1)
        izq = ttk.Frame(cuerpo)
        izq.grid(row=0, column=0, sticky="nsew")
        self._fig_en(izq, gf.fig_matriz_confusion(r))
        der = ttk.Frame(cuerpo)
        der.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ttk.Label(der, text="Reporte por segmento (muestra de prueba 20 %)").pack(anchor="w")
        t = TablaDF(der, altura=10)
        t.pack(fill="x", pady=4)
        t.mostrar(r.reporte_knn)
        leyenda = "\n".join(f"S{s}: {r.nombres_seg[s]}" for s in sorted(r.nombres_seg))
        ttk.Label(der, text=leyenda, style="Sub.TLabel", justify="left").pack(anchor="w", pady=6)

    def _tab_estacionalidad(self, r):
        fr = self.tab["estac"]
        self._vaciar(fr)
        t = TablaDF(fr, altura=min(8, len(r.calendario) + 1))
        t.pack(side="bottom", fill="x", pady=(6, 0))
        t.mostrar(r.calendario, indice=False)
        ttk.Label(fr, text="Calendario de campañas (meses pico por categoría)",
                  style="Titulo.TLabel").pack(side="bottom", anchor="w", pady=(6, 0))
        self._fig_en(fr, gf.fig_estacionalidad(r))

    def _tab_nuevo(self, r):
        fr = self.tab["nuevo"]
        self._vaciar(fr)
        ttk.Label(fr, text="Clasificar un cliente nuevo", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(fr, style="Sub.TLabel",
                  text="Ingrese los datos del cliente; el modelo KNN le asignará un segmento.").pack(anchor="w", pady=(0, 10))
        form = ttk.Frame(fr)
        form.pack(anchor="w")
        vars_ = {}

        def campo(fila, texto, defecto):
            ttk.Label(form, text=texto).grid(row=fila, column=0, sticky="w", pady=4, padx=(0, 12))
            v = tk.StringVar(value=str(defecto))
            ttk.Entry(form, textvariable=v, width=14, justify="right").grid(row=fila, column=1, pady=4)
            return v

        vars_["rec"] = campo(0, "Recencia (días desde la última compra)", 12)
        vars_["n"] = campo(1, "Número de compras", 3)
        gastos = {}
        defectos = [420, 60]
        for i, c in enumerate(r.categorias):
            gastos[c] = campo(2 + i, f"Gasto en {c} (S/)", defectos[i] if i < len(defectos) else 0)

        out = tk.Label(fr, text="", font=(FUENTE, 13, "bold"), fg=ROJO, bg=FONDO, justify="left")
        det = ttk.Label(fr, text="", style="Sub.TLabel", justify="left")

        def num(v, nombre):
            try:
                return float(v.get().replace(",", ".").strip() or 0)
            except ValueError:
                raise an.ErrorDatos(f"«{nombre}» debe ser un número.")

        def clasificar():
            try:
                rec, n = num(vars_["rec"], "Recencia"), num(vars_["n"], "Número de compras")
                g = {c: num(v, f"Gasto en {c}") for c, v in gastos.items()}
                if rec < 0 or any(x < 0 for x in g.values()):
                    raise an.ErrorDatos("Los valores no pueden ser negativos.")
                s, nombre, conf = r.clasificar_cliente(rec, n, g)
            except an.ErrorDatos as e:
                messagebox.showwarning("Datos incompletos", str(e))
                return
            out.config(text=f"Segmento asignado: S{s} · {nombre}   (confianza {conf:.0%})", fg=gf.color_segmento(s))
            p = r.perfil.loc[s]
            det.config(text=(f"Monto total del cliente: S/ {sum(g.values()):,.2f}\n\n"
                             f"Perfil promedio del segmento:  {int(p['clientes']):,} clientes ({p['%_clientes']}%)  ·  "
                             f"recencia {p['recencia_media']:.0f} días  ·  "
                             f"frecuencia {p['frecuencia_media']:.1f} compras  ·  monto S/ {p['monto_medio']:,.0f}"))

        ttk.Button(fr, text="Clasificar cliente", style="Accion.TButton", command=clasificar).pack(anchor="w", pady=14)
        out.pack(anchor="w")
        det.pack(anchor="w", pady=8)

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
    if sys.platform.startswith("win"):       # nitidez en pantallas de alta resolución
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
