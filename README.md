# Generador de Excel formateados

App Streamlit que convierte archivos de análisis en formatos listos para
importar a NetSuite.

## Módulos

| Módulo | Entrada | Salida |
|---|---|---|
| Formatear Transferencia | Excel de movimientos (cantidades +/- por tienda) + CSV de stock (opcional) | Excel de Órdenes de Transferencia, una hoja por hoja de origen |
| Formatear Orden de Pedido | `Nuevo Análisis V2.xlsx` (hojas GT, SV, HN, CR, PA) | CSV de Órdenes de Compra |
| Formatear Orden de Pedido (MX) | Igual, hojas `*-MX` (sin PA) | CSV de Órdenes de Compra |

## Cálculo de stock (opcional, solo transferencias)

Si se sube un CSV de existencias, el Excel resultante incluye cuatro columnas
adicionales: `STOCK INICIAL ORIGEN`, `STOCK NUEVO ORIGEN`,
`STOCK INICIAL DESTINO`, `STOCK NUEVO DESTINO`.

**Columnas requeridas del CSV:** `SKU`, `Ubicación del inventario`,
`Físico en ubicación`. Se aceptan con acentos correctos o con mojibake
(UTF-8 leído como Latin-1, ej. `UbicaciÃ³n`); la app repara el texto sola.

**Reglas aplicadas:**

- El número de tienda se extrae de los dígitos de `Ubicación del inventario`
  (`OD | GT | 601 TIENDA MAJADAS` → `601`).
- Las ubicaciones cuyo número no esté en `MAPEO_TIENDAS` se descartan
  (bodegas externas como `20601`, otros países). Se listan como aviso.
- El cruce con las transferencias es por **SKU**.
- Los saldos son **acumulativos**: si una tienda envía varias veces, cada
  fila descuenta del saldo corriente. Origen con 20 que envía 10 y luego 5
  termina en 5.
- **Celda de stock vacía = 0.** La tienda aparece en el archivo, por lo
  tanto su existencia es cero y puede quedar en negativo.
- **SKU+tienda ausente del archivo** = columnas en blanco, sin advertencia.
- Los saldos negativos **no bloquean** la transferencia: se generan igual y
  se listan en pantalla como error para revisión manual.
- `Nivel de stock de seguridad de la ubicación` se ignora.
- Las filas se ordenan por ID externo (ascendente, orden estable) para que
  las líneas del mismo documento queden contiguas y los saldos acumulativos
  se lean en secuencia.

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

## IDs externos: son agrupadores, no únicos

NetSuite importa las líneas agrupadas por ID externo, así que **todas las
líneas de un mismo documento comparten el mismo código**. Los IDs NO deben
ser únicos por fila. En ambos módulos las filas se ordenan por ese código
(orden estable) para que las líneas de cada documento queden contiguas.

| Módulo | Formato | Agrupa por |
|---|---|---|
| Transferencias | `OT{fecha_serial}{unidad_origen}{unidad_destino}` | par origen-destino |
| Órdenes de compra | `OcBrian{id_proveedor}{tienda}{ddmmyyyy}` | proveedor + tienda + fecha |

> Nota: ambos formatos concatenan sin separadores, por lo que en teoría
> podrían ser ambiguos. Verificado contra los catálogos actuales:
> 0 colisiones en los 166 pares de tiendas posibles dentro de una misma
> subsidiaria, y 0 en las 16.675 combinaciones proveedor × tienda.
> Conviene revalidarlo si se agregan tiendas o proveedores.

## Rendimiento

El CSV de stock puede tener cientos de miles de filas. La carga está
vectorizada con pandas y, además, filtra de entrada los SKU que no aparecen
en el archivo de movimientos.

Referencia con un CSV de 300.000 filas (30 tiendas × 15.000 SKU):

| Versión | Tiempo |
|---|---|
| Recorrido fila por fila (`iterrows`) | ~41 s |
| Vectorizado | ~2,7 s |
| Vectorizado + filtro por SKU | ~1,1 s |

El filtro solo afecta la velocidad: los saldos resultantes son idénticos
con o sin él.

## Deuda técnica conocida

- `MAPEO_TIENDAS` en `app.py` duplica datos de `data/Unidad_de_Negocio.xlsx`.
  Deberían unificarse en una sola fuente.
- El layout del archivo de entrada se detecta por posición fija de filas y
  columnas. Un cambio de estructura upstream rompe el parseo en silencio.
- Sin pruebas automatizadas.
