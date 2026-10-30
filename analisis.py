# -*- coding: utf-8 -*-
"""
Motor de análisis - Segmentación de clientes con RFM + K-Means + KNN
Amaqzoom JG S.A.C.

Este módulo contiene TODA la lógica del script original (limpieza, anonimización,
RFM, selección de k, K-Means, KNN y estacionalidad), pero sin gráficos ni
dependencias de interfaz: recibe un DataFrame y devuelve un objeto `Resultado`.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata
import warnings
from dataclasses import dataclass
from typing import Callable, Optional

import joblib
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import (accuracy_score, classification_report, confusion_matrix,
                             davies_bouldin_score, f1_score, silhouette_score)
from sklearn.model_selection import StratifiedKFold, cross_val_score, train_test_split
from sklearn.neighbors import KNeighborsClassifier
from sklearn.preprocessing import StandardScaler

warnings.filterwarnings("ignore")

# --------------------------------------------------------------------------- #
# Parámetros
# --------------------------------------------------------------------------- #
SEMILLA = 42
SALT = "amaqzoom_2026"            # sal para el hash de identificadores
K_MIN, K_MAX = 2, 8
MAX_CATEGORIAS = 6                # si hay más, las menores se agrupan en "Otros"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

COLUMNAS_REQUERIDAS = ["fecha", "id_cliente", "categoria", "monto"]
COLUMNAS_OPCIONALES = ["canal"]

# Nombres alternativos aceptados para las columnas (ya sin tildes ni mayúsculas)
ALIAS = {
    "cliente": "id_cliente", "id": "id_cliente", "idcliente": "id_cliente",
    "cod_cliente": "id_cliente", "codigo_cliente": "id_cliente",
    "fecha_venta": "fecha", "fecha_compra": "fecha",
    "linea": "categoria", "familia": "categoria", "rubro": "categoria",
    "importe": "monto", "monto_total": "monto", "total": "monto", "venta": "monto",
    "canal_venta": "canal",
}


class ErrorDatos(ValueError):
    """Error comprensible para el usuario (datos faltantes o inválidos)."""


# --------------------------------------------------------------------------- #
# Lectura y validación del archivo
# --------------------------------------------------------------------------- #
def _normalizar_columna(nombre) -> str:
    s = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode()
    s = re.sub(r"[\s\-\.]+", "_", s.strip().lower())
    s = re.sub(r"_+", "_", s).strip("_")
    return ALIAS.get(s, s)


def validar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza los nombres de columnas y comprueba que existan las requeridas."""
    original = [str(c) for c in df.columns]
    df = df.copy()
    df.columns = [_normalizar_columna(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    faltan = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltan:
        raise ErrorDatos(
            "Faltan columnas obligatorias: " + ", ".join(faltan) + ".\n\n"
            "Columnas encontradas en el archivo: " + ", ".join(original) + ".\n\n"
            "Se requieren: fecha, id_cliente, categoria, monto (y opcionalmente canal).")
    df = df.dropna(how="all")
    if df.empty:
        raise ErrorDatos("El archivo no contiene filas con datos.")
    return df.reset_index(drop=True)


def listar_hojas(ruta: str) -> list[str]:
    if os.path.splitext(ruta)[1].lower() in (".csv", ".txt"):
        return ["(CSV)"]
    return pd.ExcelFile(ruta).sheet_names


def leer_datos(ruta: str, hoja: Optional[str] = None) -> pd.DataFrame:
    """Lee un Excel (.xlsx/.xlsm/.xls) o CSV y devuelve el DataFrame validado."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".csv", ".txt"):
        try:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(ruta, sep=None, engine="python", encoding="latin-1")
    elif ext in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(ruta, sheet_name=hoja if hoja not in (None, "(CSV)") else 0)
    else:
        raise ErrorDatos(f"Formato no soportado: {ext}. Use .xlsx, .xls o .csv")
    return validar_columnas(df)


# --------------------------------------------------------------------------- #
# Datos sintéticos de demostración (idénticos al script original)
# --------------------------------------------------------------------------- #
def generar_datos_sinteticos(n_clientes: int = 1200, semilla: int = SEMILLA) -> pd.DataFrame:
    """Genera ventas de 12 meses (oct-2025 a sep-2026) con 'suciedad' típica de Excel."""
    categorias = ["Ventiladores", "Cocina", "Audio"]
    rng = np.random.default_rng(semilla)
    meses = pd.period_range("2025-10", "2026-09", freq="M")
    perfiles = {
        "vent":  dict(p=.33, lam=1.6, cat=[.85, .08, .07], mult=1.0, peso={12: 5, 1: 5, 2: 5, 3: 4}),
        "coc":   dict(p=.25, lam=2.4, cat=[.07, .86, .07], mult=0.9, peso={10: 2, 11: 3, 12: 4, 5: 4}),
        "aud":   dict(p=.20, lam=1.5, cat=[.06, .08, .86], mult=0.8, peso={11: 5, 12: 4, 7: 3}),
        "vip":   dict(p=.10, lam=9.0, cat=[.34, .33, .33], mult=2.2, peso={}),
        "inact": dict(p=.12, lam=0.6, cat=[.34, .33, .33], mult=1.0, peso={10: 3, 11: 3, 12: 2, 1: 1}),
    }
    nombres = list(perfiles)
    prob = [perfiles[n]["p"] for n in nombres]
    ticket_base = {"Ventiladores": 150, "Cocina": 120, "Audio": 90}
    canales = ["Tienda física", "Sodimac", "Saga Falabella", "Plaza Vea"]
    filas = []
    for _ in range(n_clientes):
        cid = "9" + str(rng.integers(10**7, 10**8 - 1))
        pf = perfiles[rng.choice(nombres, p=prob)]
        n_tx = 1 + rng.poisson(pf["lam"])
        pesos = np.array([pf["peso"].get(m.month, 1) for m in meses], float)
        pesos = pesos / pesos.sum()
        for _ in range(n_tx):
            per = meses[rng.choice(len(meses), p=pesos)]
            fecha = pd.Timestamp(per.year, per.month, int(rng.integers(1, 28)))
            cat = rng.choice(categorias, p=pf["cat"])
            monto = rng.lognormal(np.log(ticket_base[cat] * pf["mult"]), 0.45)
            filas.append([fecha.strftime("%d/%m/%Y"), cid,
                          rng.choice(canales, p=[.45, .25, .2, .1]), cat, round(float(monto), 2)])
    df = pd.DataFrame(filas, columns=["fecha", "id_cliente", "canal", "categoria", "monto"])
    # --- ensuciar los datos como en un Excel real ---
    df["monto"] = df["monto"].astype(object)
    m = rng.random(len(df)) < .30
    df.loc[m, "monto"] = df.loc[m, "monto"].apply(lambda x: f"S/ {x:,.2f}")
    df.loc[rng.random(len(df)) < .01, "monto"] = np.nan
    df.loc[rng.random(len(df)) < .005, "id_cliente"] = np.nan
    df.loc[rng.random(len(df)) < .10, "categoria"] = df["categoria"].str.upper() + " "
    df = pd.concat([df, df.sample(frac=.02, random_state=semilla)], ignore_index=True)
    return df.sample(frac=1, random_state=semilla).reset_index(drop=True)


def plantilla_ejemplo() -> pd.DataFrame:
    """Pocas filas de ejemplo para que el usuario vea el formato esperado."""
    return pd.DataFrame({
        "fecha": ["05/10/2025", "18/11/2025", "02/12/2025", "14/01/2026", "20/02/2026", "09/03/2026"],
        "id_cliente": ["912345678", "912345678", "987654321", "955512345", "987654321", "955512345"],
        "canal": ["Tienda física", "Sodimac", "Saga Falabella", "Plaza Vea", "Tienda física", "Sodimac"],
        "categoria": ["Ventiladores", "Cocina", "Audio", "Ventiladores", "Audio", "Cocina"],
        "monto": [189.90, "S/ 120.50", 95.00, 210.00, "S/ 88.00", 134.50],
    })


# --------------------------------------------------------------------------- #
# Limpieza y anonimización (HU02)
# --------------------------------------------------------------------------- #
def _id_a_texto(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (float, np.floating)) and float(x).is_integer():
        return str(int(x))
    return str(x).strip()


def _parsear_fechas(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    try:
        return pd.to_datetime(serie, dayfirst=True, errors="coerce", format="mixed")
    except (TypeError, ValueError):      # pandas < 2.0
        return pd.to_datetime(serie, dayfirst=True, errors="coerce")


def limpiar(df: pd.DataFrame, anonimizar: bool = True):
    informe = {"registros_iniciales": len(df)}
    df = df.drop_duplicates()
    informe["duplicados_eliminados"] = informe["registros_iniciales"] - len(df)

    df = df.copy()
    df["id_cliente"] = df["id_cliente"].map(_id_a_texto)
    n0 = len(df)
    df = df.dropna(subset=["id_cliente", "monto", "fecha"]).copy()
    informe["nulos_eliminados"] = n0 - len(df)

    df["monto"] = pd.to_numeric(
        df["monto"].astype(str).str.replace(r"[^\d.\-]", "", regex=True).replace("", np.nan),
        errors="coerce")
    df["fecha"] = _parsear_fechas(df["fecha"])
    df["categoria"] = (df["categoria"].fillna("Sin categoría").astype(str)
                       .str.strip().str.title().replace("", "Sin Categoría"))
    n1 = len(df)
    df = df.dropna(subset=["monto", "fecha"])
    df = df[df["monto"] > 0]
    informe["montos_fechas_invalidos"] = n1 - len(df)

    if anonimizar:   # hash SHA-256 con sal: el identificador original no se conserva
        df["id_cliente"] = df["id_cliente"].apply(
            lambda x: hashlib.sha256((SALT + x).encode()).hexdigest()[:12])
    informe["registros_finales"] = len(df)
    informe["clientes_unicos"] = int(df["id_cliente"].nunique())
    return df.reset_index(drop=True), informe


# --------------------------------------------------------------------------- #
# Resultado
# --------------------------------------------------------------------------- #
def preparar_features(df: pd.DataFrame, cols_modelo: list[str]) -> pd.DataFrame:
    """log1p a frecuencia y monto (distribuciones sesgadas); recencia y % sin transformar."""
    X = df[cols_modelo].copy()
    X["frecuencia"] = np.log1p(X["frecuencia"])
    X["monto"] = np.log1p(X["monto"])
    return X


@dataclass
class Resultado:
    ventas: pd.DataFrame
    informe: dict
    categorias: list
    clientes: pd.DataFrame
    cols_pct: list
    cols_modelo: list
    tabla_k: pd.DataFrame
    k_opt: int
    k_automatico: bool
    silhouette: float
    davies_bouldin: float
    perfil: pd.DataFrame
    nombres_seg: dict
    proyeccion: np.ndarray
    var_pca: float
    cv_knn: dict
    k_knn: int
    exactitud: float
    f1_macro: float
    reporte_knn: pd.DataFrame
    matriz_conf: np.ndarray
    etiquetas: list
    est_idx: pd.DataFrame
    calendario: pd.DataFrame
    resumen: dict
    anonimizado: bool
    scaler: StandardScaler
    kmeans: KMeans
    knn: KNeighborsClassifier

    def clasificar_cliente(self, recencia: float, n_compras: float, gastos: dict):
        """Clasifica un cliente nuevo (HU06). `gastos` = {categoria: monto gastado}."""
        g = np.array([float(gastos.get(c, 0.0)) for c in self.categorias], float)
        if g.sum() <= 0:
            raise ErrorDatos("Ingrese un gasto mayor a cero en al menos una categoría.")
        if n_compras < 1:
            raise ErrorDatos("El número de compras debe ser al menos 1.")
        fila = pd.DataFrame([[recencia, n_compras, g.sum(), *(g / g.sum())]], columns=self.cols_modelo)
        xs = self.scaler.transform(preparar_features(fila, self.cols_modelo))
        s = int(self.knn.predict(xs)[0])
        conf = float(self.knn.predict_proba(xs)[0].max())
        return s, self.nombres_seg[s], conf


# --------------------------------------------------------------------------- #
# Análisis completo
# --------------------------------------------------------------------------- #
def _nombrar(fila, ref, cols_pct, categorias):
    """VIP si frecuencia y monto están en el 25 % superior; si no, categoría dominante."""
    dom = categorias[int(np.argmax([fila[c] for c in cols_pct]))]
    es_vip = (fila["frecuencia_media"] >= ref["frecuencia_media"].quantile(.75) and
              fila["monto_medio"] >= ref["monto_medio"].quantile(.75))
    if es_vip:
        return "VIP multicategoría" if max(fila[c] for c in cols_pct) < .6 else f"VIP {dom}"
    return f"Comprador de {dom}"


def ejecutar_analisis(df_raw: pd.DataFrame, anonimizar: bool = True,
                      k_manual: Optional[int] = None,
                      progreso: Callable[[str], None] = lambda m: None) -> Resultado:
    # 1. Limpieza ---------------------------------------------------------- #
    progreso("Preparando y revisando tus ventas…")
    ventas, informe = limpiar(validar_columnas(df_raw), anonimizar)
    if informe["clientes_unicos"] < 10:
        raise ErrorDatos(
            f"Tras la limpieza solo quedaron {informe['clientes_unicos']} clientes únicos. "
            "Se necesitan al menos 10 para segmentar. Revise fechas, montos e id_cliente.")

    # Categorías: las de mayor gasto; el resto se agrupa en "Otros"
    tot = ventas.groupby("categoria")["monto"].sum().sort_values(ascending=False)
    if len(tot) > MAX_CATEGORIAS:
        top = tot.index[:MAX_CATEGORIAS].tolist()
        ventas["categoria"] = ventas["categoria"].where(ventas["categoria"].isin(top), "Otros")
        categorias = top + ([] if "Otros" in top else ["Otros"])
    else:
        categorias = tot.index.tolist()

    # 2. RFM + % de gasto por categoría (HU03) ----------------------------- #
    progreso("Revisando los hábitos de compra de tus clientes…")
    fecha_ref = ventas["fecha"].max() + pd.Timedelta(days=1)
    rfm = ventas.groupby("id_cliente").agg(
        recencia=("fecha", lambda s: (fecha_ref - s.max()).days),
        frecuencia=("fecha", "count"),
        monto=("monto", "sum"))
    gasto_cat = (ventas.pivot_table(index="id_cliente", columns="categoria", values="monto",
                                    aggfunc="sum", fill_value=0)
                 .reindex(columns=categorias, fill_value=0))
    pct = gasto_cat.div(gasto_cat.sum(axis=1), axis=0)
    cols_pct = [f"pct_{c.lower()}" for c in categorias]
    pct.columns = cols_pct
    clientes = rfm.join(pct).round(3)
    cols_modelo = ["recencia", "frecuencia", "monto"] + cols_pct

    scaler = StandardScaler().fit(preparar_features(clientes, cols_modelo))
    X = scaler.transform(preparar_features(clientes, cols_modelo))

    # 3. Selección de k (HU05) --------------------------------------------- #
    progreso("Comparando los perfiles de compra…")
    k_max = min(K_MAX, len(clientes) - 1)
    filas_k = []
    for k in range(K_MIN, k_max + 1):
        km = KMeans(n_clusters=k, n_init=10, random_state=SEMILLA).fit(X)
        filas_k.append(dict(k=k, inercia=km.inertia_,
                            silhouette=silhouette_score(X, km.labels_),
                            davies_bouldin=davies_bouldin_score(X, km.labels_)))
    tabla_k = pd.DataFrame(filas_k)
    k_auto = k_manual is None
    if k_auto:
        k_opt = int(tabla_k.loc[tabla_k["silhouette"].idxmax(), "k"])
    else:
        k_opt = int(min(max(k_manual, K_MIN), k_max))

    # 4. K-Means y validación (HU04) --------------------------------------- #
    progreso("Organizando los clientes en perfiles de compra…")
    kmeans = KMeans(n_clusters=k_opt, n_init=10, random_state=SEMILLA).fit(X)
    clientes["segmento"] = kmeans.labels_
    sil = float(silhouette_score(X, kmeans.labels_))
    dbi = float(davies_bouldin_score(X, kmeans.labels_))

    perfil = clientes.groupby("segmento").agg(
        clientes=("recencia", "size"), recencia_media=("recencia", "mean"),
        frecuencia_media=("frecuencia", "mean"), monto_medio=("monto", "mean"),
        **{c: (c, "mean") for c in cols_pct}).round(2)
    perfil["%_clientes"] = (100 * perfil["clientes"] / len(clientes)).round(1)
    perfil["nombre"] = perfil.apply(_nombrar, axis=1, ref=perfil, cols_pct=cols_pct, categorias=categorias)
    dup = perfil["nombre"].duplicated(keep=False)          # evitar nombres repetidos
    perfil.loc[dup, "nombre"] = [f"{n} (S{s})" for s, n in perfil.loc[dup, "nombre"].items()]
    nombres_seg = perfil["nombre"].to_dict()
    clientes["nombre_segmento"] = clientes["segmento"].map(nombres_seg)

    pca = PCA(n_components=2, random_state=SEMILLA).fit(X)
    proyeccion = pca.transform(X)
    var_pca = float(pca.explained_variance_ratio_.sum())

    # 5. Clasificación KNN (HU06) ------------------------------------------ #
    progreso("Preparando información para conocer los perfiles de clientes…")
    y = clientes["segmento"].values
    conteo = np.bincount(y)
    estrat = y if conteo[conteo > 0].min() >= 2 else None
    try:
        X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=.20, stratify=estrat, random_state=SEMILLA)
    except ValueError:
        X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=.20, random_state=SEMILLA)
    conteo_tr = np.bincount(y_tr)
    n_splits = int(min(5, conteo_tr[conteo_tr > 0].min()))
    candidatos = [k for k in (3, 5, 7, 9, 11, 15) if k <= max(1, len(y_tr) // 2)] or [1]
    if n_splits >= 2:
        cv = StratifiedKFold(n_splits, shuffle=True, random_state=SEMILLA)
        cv_knn = {k: float(cross_val_score(KNeighborsClassifier(n_neighbors=k), X_tr, y_tr, cv=cv).mean())
                  for k in candidatos}
        k_knn = max(cv_knn, key=cv_knn.get)
    else:
        cv_knn, k_knn = {}, min(5, len(y_tr))
    knn_prueba = KNeighborsClassifier(n_neighbors=k_knn).fit(X_tr, y_tr)
    y_pred = knn_prueba.predict(X_te)
    exactitud = float(accuracy_score(y_te, y_pred))
    f1m = float(f1_score(y_te, y_pred, average="macro"))
    etiquetas = sorted(set(y))
    nombres_cls = [f"S{s}" for s in etiquetas]
    rep = classification_report(y_te, y_pred, labels=etiquetas, target_names=nombres_cls,
                                output_dict=True, zero_division=0)
    rep.pop("accuracy", None)
    reporte = pd.DataFrame(rep).T.round(3)
    reporte.columns = ["precisión", "sensibilidad (recall)", "F1", "soporte"]
    reporte["soporte"] = reporte["soporte"].astype(int)
    reporte.index.name = "segmento"
    cm = confusion_matrix(y_te, y_pred, labels=etiquetas)
    knn = KNeighborsClassifier(n_neighbors=k_knn).fit(X, y)     # modelo final con todos los datos

    # 6. Estacionalidad y calendario (HU08) -------------------------------- #
    progreso("Identificando las temporadas con más ventas…")
    mes = ventas["fecha"].dt.month
    est = (ventas.assign(mes=mes).pivot_table(index="mes", columns="categoria", values="monto",
                                              aggfunc="sum", fill_value=0)
           .reindex(range(1, 13), fill_value=0))
    est_idx = est / est.mean()                                   # 1.0 = mes promedio
    cal = []
    for c in est_idx.columns:
        top3 = est_idx[c].nlargest(3).index.tolist()
        cal.append({"categoria": c, "meses_pico": ", ".join(MESES[m - 1] for m in sorted(top3)),
                    "indice_pico": round(float(est_idx[c].max()), 2),
                    "accion": "Campaña 3-4 semanas antes del pico y refuerzo de stock"})
    calendario = pd.DataFrame(cal)

    resumen = dict(clientes=int(len(clientes)), k=k_opt, silhouette=round(sil, 3),
                   davies_bouldin=round(dbi, 3), knn_K=int(k_knn),
                   exactitud=round(exactitud, 3), f1_macro=round(f1m, 3))
    progreso("Análisis de ventas completado.")
    return Resultado(
        ventas=ventas, informe=informe, categorias=categorias, clientes=clientes,
        cols_pct=cols_pct, cols_modelo=cols_modelo, tabla_k=tabla_k, k_opt=k_opt,
        k_automatico=k_auto, silhouette=sil, davies_bouldin=dbi, perfil=perfil,
        nombres_seg=nombres_seg, proyeccion=proyeccion, var_pca=var_pca, cv_knn=cv_knn,
        k_knn=int(k_knn), exactitud=exactitud, f1_macro=f1m, reporte_knn=reporte,
        matriz_conf=cm, etiquetas=etiquetas, est_idx=est_idx, calendario=calendario,
        resumen=resumen, anonimizado=anonimizar, scaler=scaler, kmeans=kmeans, knn=knn)


# --------------------------------------------------------------------------- #
# Exportación
# --------------------------------------------------------------------------- #
def _autoajustar(ws):
    for col in ws.columns:
        ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(10, ancho + 2), 50)


def exportar_tablas(res: Resultado, carpeta: str) -> str:
    """Guarda Excel con todas las tablas + dataset limpio + modelo + resumen JSON."""
    os.makedirs(carpeta, exist_ok=True)
    ruta_xlsx = os.path.join(carpeta, "resultados_segmentacion.xlsx")
    with pd.ExcelWriter(ruta_xlsx, engine="openpyxl") as w:
        pd.DataFrame(list(res.resumen.items()), columns=["métrica", "valor"]).to_excel(
            w, sheet_name="Resumen", index=False)
        pd.DataFrame(list(res.informe.items()), columns=["concepto", "valor"]).to_excel(
            w, sheet_name="Limpieza", index=False)
        res.perfil.to_excel(w, sheet_name="Perfiles_segmentos")
        res.clientes.to_excel(w, sheet_name="Clientes_segmentados")
        res.tabla_k.round(4).to_excel(w, sheet_name="Seleccion_k", index=False)
        res.reporte_knn.to_excel(w, sheet_name="KNN_reporte")
        res.est_idx.round(3).rename_axis("mes").to_excel(w, sheet_name="Estacionalidad")
        res.calendario.to_excel(w, sheet_name="Calendario", index=False)
        for ws in w.book.worksheets:
            _autoajustar(ws)
    res.ventas.to_csv(os.path.join(carpeta, "dataset_limpio.csv"), index=False, encoding="utf-8-sig")
    with open(os.path.join(carpeta, "resumen_metricas.json"), "w", encoding="utf-8") as f:
        json.dump(res.resumen, f, indent=2, ensure_ascii=False)
    joblib.dump({"scaler": res.scaler, "kmeans": res.kmeans, "knn": res.knn,
                 "nombres": res.nombres_seg, "cols": res.cols_modelo,
                 "categorias": res.categorias},
                os.path.join(carpeta, "modelo_segmentacion.joblib"))
    return ruta_xlsx
