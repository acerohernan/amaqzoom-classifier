# Amaqzoom · Análisis de ventas (aplicación de escritorio)

Aplicación de escritorio para revisar información de ventas y clientes a partir de
un Excel o CSV. Presenta indicadores de ventas, tipos de clientes, oportunidades de
retorno y temporadas de mayor venta en lenguaje de negocio.

## Instalación (una sola vez)
Requiere **Python 3.9 o superior** (en Windows, marque «Add Python to PATH» al instalarlo).

    pip install -r requirements.txt

En Linux, además: `sudo apt install python3-tk`.

## Uso
    python app.py          (en Windows también: doble clic en iniciar_app.bat)

1. **Importar ventas (Excel / CSV)…** (o «Datos de demostración» para probar).
2. Elija la hoja si el libro tiene varias.
3. Decida si desea proteger los identificadores de clientes durante el análisis.
4. **Analizar mis ventas** y revise el resumen, los perfiles de clientes y las temporadas.
5. En **Clasificar nuevo cliente**, ingrese los días desde su última compra, el número de
   compras y el gasto por categoría para estimar su perfil con KNN.
6. **Exportar resultados…** guarda un Excel con todas las tablas, los gráficos en PNG,
   el dataset limpio, el resumen JSON y el modelo entrenado (`.joblib`).

## Formato del Excel
Una fila por transacción, con estas columnas (no importan mayúsculas ni tildes):

| fecha | cliente | canal (opcional) | categoria | monto |
|---|---|---|---|---|
| 05/10/2025 | 912345678 | Tienda física | Ventiladores | S/ 189.90 |

Botón **Descargar plantilla** genera un archivo de ejemplo. También se aceptan nombres
alternativos como `id_cliente`, `id`, `importe`, `total`, `línea` o `familia`.
La suciedad típica (duplicados, nulos, «S/» en montos, categorías en mayúsculas) se limpia sola.
Si hay más de 6 categorías, las de menor gasto se agrupan en «Otros».

## Estructura
- `analisis.py`  → preparación de datos y análisis de ventas/clientes.
- `graficos.py`  → gráficos de ventas, perfiles de clientes y temporadas (pantalla y PNG).
- `app.py`       → interfaz tkinter.

## Generar un .exe (opcional)
    pip install pyinstaller
    pyinstaller --noconfirm --windowed --name AmaqzoomSegmentacion app.py

El ejecutable queda en `dist/AmaqzoomSegmentacion/`.
