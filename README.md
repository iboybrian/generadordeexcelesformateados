# Generador de Excel formateados

App Streamlit que convierte archivos de análisis en formatos listos para
importar a NetSuite.

## Módulos

| Módulo | Entrada | Salida |
|---|---|---|
| Formatear Transferencia | Excel de movimientos (cantidades +/- por tienda) | Excel de Órdenes de Transferencia, una hoja por hoja de origen |
| Formatear Orden de Pedido | `Nuevo Análisis V2.xlsx` (hojas GT, SV, HN, CR, PA) | CSV de Órdenes de Compra |
| Formatear Orden de Pedido (MX) | Igual, hojas `*-MX` (sin PA) | CSV de Órdenes de Compra |

## Instalación

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura esperada del archivo de entrada

**Transferencias:** columnas `SKU` e `ID interno`, más una columna por tienda
cuyo encabezado empiece con el número de tienda (`601`, `602 TIENDA X`).
Las cantidades negativas son origen, las positivas destino. Cada SKU debe
balancear: la suma de salidas debe igualar la suma de entradas.

**Órdenes de compra:** encabezados de tienda en la fila 13 (índice 12), datos
desde la fila 14. Columnas fijas por posición: proveedor (A), ID interno (B),
SKU (C), descripción (D), pack (E). Estas posiciones están definidas como
constantes al inicio de `app.py`.

## Catálogos

`data/Proveedores.xlsx` — columnas `ID_PROVEEDOR`, `Proveedor`
`data/Unidad_de_Negocio.xlsx` — columnas `No. Tienda`, `UNIDAD DE NEGOCIO`,
`CENTRO DE COSTO`, `SUBSIDARIA`, `Unidad de Negocio del inventario`

Los proveedores que no coincidan con el catálogo aparecen como `#SIN_MATCH`
y se listan como error en pantalla antes de descargar.

## Deuda técnica conocida

- `MAPEO_TIENDAS` en `app.py` duplica datos de `data/Unidad_de_Negocio.xlsx`.
  Deberían unificarse en una sola fuente.
- El layout del archivo de entrada se detecta por posición fija de filas y
  columnas. Un cambio de estructura upstream rompe el parseo en silencio.
- Sin pruebas automatizadas.
