# -*- coding: utf-8 -*-
"""Gráficos del análisis. Cada función recibe un `Resultado` y devuelve una `Figure`."""
import os

import numpy as np
from matplotlib.figure import Figure

from analisis import MESES, Resultado

PALETA = ["#B50E2E", "#1F7A8C", "#E8A33D", "#4C6EF5", "#7B9E3C", "#8E44AD", "#555555", "#E67E22"]


def color_segmento(s) -> str:
    return PALETA[int(s) % len(PALETA)]


def _fig(ancho=8, alto=4.5) -> Figure:
    return Figure(figsize=(ancho, alto), dpi=100, layout="constrained")


def _etq(res: Resultado, s) -> str:
    return nombre_cliente(res.nombres_seg[s])


def nombre_cliente(nombre: str) -> str:
    if nombre == "VIP multicategoría":
        return "Clientes de alto valor · varias categorías"
    if nombre.startswith("VIP "):
        return f"Clientes de alto valor · {nombre[4:]}"
    if nombre.startswith("Comprador de "):
        return f"Compradores de {nombre[len('Comprador de '):]}"
    return nombre


def fig_perfil(res: Resultado) -> Figure:
    fig = _fig(11, 5.6)
    ax = fig.subplots(2, 2)
    p = res.perfil
    etq = [nombre_cliente(str(p.at[s, "nombre"])) for s in p.index]
    cols = [color_segmento(s) for s in p.index]
    ax[0, 0].bar(etq, p["recencia_media"].to_numpy(dtype=float), color=cols)
    ax[0, 0].set_title("Días promedio desde la última compra")
    ax[0, 0].set_ylabel("Días (menos días = compra más reciente)")
    ax[0, 1].bar(etq, p["frecuencia_media"].to_numpy(dtype=float), color=cols)
    ax[0, 1].set_title("Compras promedio por cliente")
    ax[0, 1].set_ylabel("Compras")
    ax[1, 0].bar(etq, p["monto_medio"].to_numpy(dtype=float), color=cols)
    ax[1, 0].set_title("Ventas registradas por cliente")
    ax[1, 0].set_ylabel("Soles (S/)")
    base = np.zeros(len(p))
    for i, (c, cat) in enumerate(zip(res.cols_pct, res.categorias)):
        vals = 100 * p[c].to_numpy(dtype=float)
        ax[1, 1].bar(etq, vals, bottom=base, label=cat, color=PALETA[(i + 1) % len(PALETA)])
        base += vals
    ax[1, 1].set_title("Distribución del gasto por categoría")
    ax[1, 1].set_ylabel("% del gasto")
    ax[1, 1].set_ylim(0, 100)
    ax[1, 1].legend(fontsize=7, loc="upper center", ncol=3, frameon=False)
    for a in ax.flat:
        a.tick_params(axis="x", labelrotation=15, labelsize=7)
    fig.suptitle("Hábitos de compra por tipo de cliente")
    return fig


def fig_tamano_segmentos(res: Resultado) -> Figure:
    fig = _fig(6, 3.6)
    ax = fig.subplots()
    p = res.perfil
    etq = [_etq(res, s) for s in p.index]
    barras = ax.barh(etq[::-1], p["clientes"].to_numpy(dtype=float)[::-1],
                     color=[color_segmento(s) for s in p.index][::-1])
    for b, pc in zip(barras, p["%_clientes"].to_numpy(dtype=float)[::-1]):
        ax.text(b.get_width(), b.get_y() + b.get_height() / 2, f" {pc:.0f}%", va="center", fontsize=8)
    ax.set(title="¿Cómo se distribuyen tus clientes?", xlabel="Número de clientes")
    ax.margins(x=.15)
    return fig


def fig_estacionalidad(res: Resultado) -> Figure:
    fig = _fig(9, 4.2)
    ax = fig.subplots()
    for i, c in enumerate(res.est_idx.columns):
        ax.plot(MESES, res.est_idx[c].to_numpy(dtype=float) * 100, "o-", label=c,
                color=PALETA[(i + 1) % len(PALETA)])
    ax.axhline(100, color="gray", ls="--", lw=.8, label="Promedio mensual")
    ax.set(title="Ventas por mes y categoría",
           ylabel="Ventas comparadas con el promedio mensual (%)")
    ax.legend(fontsize=8)
    return fig


FIGURAS = {
    "fig_perfil_segmentos.png": fig_perfil,
    "fig_tamano_segmentos.png": fig_tamano_segmentos,
    "fig_estacionalidad.png": fig_estacionalidad,
}


def guardar_figuras(res: Resultado, carpeta: str) -> None:
    os.makedirs(carpeta, exist_ok=True)
    for nombre, f in FIGURAS.items():
        f(res).savefig(os.path.join(carpeta, nombre), dpi=150)
