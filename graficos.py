# -*- coding: utf-8 -*-
"""Gráficos del análisis. Cada función recibe un `Resultado` y devuelve una `Figure`."""
import os

import numpy as np
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator
from sklearn.metrics import ConfusionMatrixDisplay

from analisis import MESES, Resultado

PALETA = ["#B50E2E", "#1F7A8C", "#E8A33D", "#4C6EF5", "#7B9E3C", "#8E44AD", "#555555", "#E67E22"]


def color_segmento(s) -> str:
    return PALETA[int(s) % len(PALETA)]


def _fig(ancho=8, alto=4.5) -> Figure:
    return Figure(figsize=(ancho, alto), dpi=100, layout="constrained")


def _etq(res: Resultado, s) -> str:
    return f"S{s}: {res.nombres_seg[s]}"


def fig_codo_silhouette(res: Resultado) -> Figure:
    fig = _fig(10, 3.8)
    ax = fig.subplots(1, 2)
    t = res.tabla_k
    ax[0].plot(t["k"], t["inercia"], "o-", color="#B50E2E")
    ax[0].axvline(res.k_opt, ls="--", color="gray")
    ax[0].set(title="Método del codo", xlabel="Número de grupos (k)", ylabel="Inercia")
    ax[1].plot(t["k"], t["silhouette"], "o-", color="#1F7A8C")
    ax[1].axvline(res.k_opt, ls="--", color="gray", label=f"k elegido = {res.k_opt}")
    ax[1].set(title="Índice de silhouette", xlabel="Número de grupos (k)", ylabel="Silhouette")
    ax[1].legend(fontsize=8)
    for a in ax:
        a.xaxis.set_major_locator(MaxNLocator(integer=True))
    return fig


def fig_pca(res: Resultado) -> Figure:
    fig = _fig(7.5, 5.2)
    ax = fig.subplots()
    seg = res.clientes["segmento"].values
    for s in sorted(set(seg)):
        m = seg == s
        ax.scatter(res.proyeccion[m, 0], res.proyeccion[m, 1], s=12, alpha=.6,
                   color=color_segmento(s), label=_etq(res, s))
    ax.set(title=f"Segmentos K-Means proyectados en 2D (PCA, {100 * res.var_pca:.0f}% de varianza)",
           xlabel="Componente 1", ylabel="Componente 2")
    ax.legend(fontsize=8)
    return fig


def fig_perfil(res: Resultado) -> Figure:
    fig = _fig(11, 5.6)
    ax = fig.subplots(2, 2)
    p = res.perfil
    etq = [f"S{s}" for s in p.index]
    cols = [color_segmento(s) for s in p.index]
    ax[0, 0].bar(etq, p["recencia_media"], color=cols); ax[0, 0].set_title("Recencia media (días)")
    ax[0, 1].bar(etq, p["frecuencia_media"], color=cols); ax[0, 1].set_title("Frecuencia media (compras)")
    ax[1, 0].bar(etq, p["monto_medio"], color=cols); ax[1, 0].set_title("Monto medio (S/)")
    base = np.zeros(len(p))
    for i, (c, cat) in enumerate(zip(res.cols_pct, res.categorias)):
        vals = 100 * p[c].values
        ax[1, 1].bar(etq, vals, bottom=base, label=cat, color=PALETA[(i + 1) % len(PALETA)])
        base += vals
    ax[1, 1].set_title("Mix de gasto por categoría (%)")
    ax[1, 1].set_ylim(0, 128)
    ax[1, 1].legend(fontsize=7, loc="upper center", ncol=3, frameon=False)
    fig.suptitle("Perfil RFM por segmento")
    return fig


def fig_tamano_segmentos(res: Resultado) -> Figure:
    fig = _fig(6, 3.6)
    ax = fig.subplots()
    p = res.perfil
    etq = [_etq(res, s) for s in p.index]
    barras = ax.barh(etq[::-1], p["clientes"].values[::-1],
                     color=[color_segmento(s) for s in p.index][::-1])
    for b, pc in zip(barras, p["%_clientes"].values[::-1]):
        ax.text(b.get_width(), b.get_y() + b.get_height() / 2, f" {pc:.0f}%", va="center", fontsize=8)
    ax.set(title="Clientes por segmento", xlabel="N.º de clientes")
    ax.margins(x=.15)
    return fig


def fig_matriz_confusion(res: Resultado) -> Figure:
    fig = _fig(5.2, 4.6)
    ax = fig.subplots()
    ConfusionMatrixDisplay(res.matriz_conf, display_labels=[f"S{s}" for s in res.etiquetas]).plot(
        ax=ax, cmap="Reds", colorbar=False)
    ax.set_title("Matriz de confusión · KNN\n(muestra de prueba 20 %)", fontsize=10)
    return fig


def fig_estacionalidad(res: Resultado) -> Figure:
    fig = _fig(9, 4.2)
    ax = fig.subplots()
    for i, c in enumerate(res.est_idx.columns):
        ax.plot(MESES, res.est_idx[c].values, "o-", label=c, color=PALETA[(i + 1) % len(PALETA)])
    ax.axhline(1, color="gray", ls="--", lw=.8)
    ax.set(title="Estacionalidad de ventas por categoría (índice; 1 = mes promedio)",
           ylabel="Índice estacional")
    ax.legend(fontsize=8)
    return fig


FIGURAS = {
    "fig_codo_silhouette.png": fig_codo_silhouette,
    "fig_pca_segmentos.png": fig_pca,
    "fig_perfil_segmentos.png": fig_perfil,
    "fig_tamano_segmentos.png": fig_tamano_segmentos,
    "fig_matriz_confusion.png": fig_matriz_confusion,
    "fig_estacionalidad.png": fig_estacionalidad,
}


def guardar_figuras(res: Resultado, carpeta: str) -> None:
    os.makedirs(carpeta, exist_ok=True)
    for nombre, f in FIGURAS.items():
        f(res).savefig(os.path.join(carpeta, nombre), dpi=150)
