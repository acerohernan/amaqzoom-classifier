# Amaqzoom · Segmentación de clientes (aplicación de escritorio)

Convierte `poc_segmentacion_amaqzoom.py` en una aplicación con interfaz gráfica:
carga un Excel/CSV de ventas, ejecuta el análisis (limpieza → RFM → K-Means → KNN →
estacionalidad) y muestra los resultados en pestañas.

## Instalación (una sola vez)
Requiere **Python 3.9 o superior** (en Windows, marque «Add Python to PATH» al instalarlo).

    pip install -r requirements.txt

En Linux, además: `sudo apt install python3-tk`.

## Uso
    python app.py          (en Windows también: doble clic en iniciar_app.bat)

1. **Cargar Excel / CSV…** (o «Datos de demostración» para probar).
2. Elija la hoja si el libro tiene varias.
3. Opcional: desmarque la anonimización o fije el número de grupos (k).
4. **▶ Ejecutar análisis** y revise las pestañas.
5. **Exportar resultados…** guarda un Excel con todas las tablas, los gráficos en PNG,
   el dataset limpio, el resumen JSON y el modelo entrenado (`.joblib`).

## Formato del Excel
Una fila por transacción, con estas columnas (no importan mayúsculas ni tildes):

| fecha | id_cliente | canal (opcional) | categoria | monto |
|---|---|---|---|---|
| 05/10/2025 | 912345678 | Tienda física | Ventiladores | S/ 189.90 |

Botón **Descargar plantilla** genera un archivo de ejemplo. La suciedad típica
(duplicados, nulos, «S/» en montos, categorías en mayúsculas) se limpia sola.
Si hay más de 6 categorías, las de menor gasto se agrupan en «Otros».

## Estructura
- `analisis.py`  → toda la lógica del script original, sin gráficos (reutilizable).
- `graficos.py`  → figuras de matplotlib (pantalla y PNG).
- `app.py`       → interfaz tkinter.

## Generar un .exe (opcional)
    pip install pyinstaller
    pyinstaller --noconfirm --windowed --name AmaqzoomSegmentacion app.py

El ejecutable queda en `dist/AmaqzoomSegmentacion/`.
