"""
Generador de documentos formateados para NetSuite.

Tres utilidades:
  1. Formatear Transferencia      -> Excel de Ordenes de Transferencia (OT)
  2. Formatear Orden de Pedido    -> CSV de Ordenes de Compra
  3. Formatear Orden de Pedido MX -> igual, con hojas MX

Ejecutar:  streamlit run app.py
"""

import csv
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

# ------------------------------------------------------------
# CONSTANTES DE LAYOUT DEL ARCHIVO DE ENTRADA
# Si el archivo fuente cambia de estructura, ajustar aqui.
# ------------------------------------------------------------
FILA_ENCABEZADOS = 12       # fila (0-indexed) con los numeros de tienda
FILA_PRIMER_DATO = 13       # primera fila con datos
COL_PROVEEDOR = 0
COL_ID_INTERNO = 1
COL_SKU = 2
COL_DESCRIPCION = 3
COL_PACK = 4

TOLERANCIA = 1e-9           # comparaciones de punto flotante

# Rutas de catalogos, relativas a este archivo (no al cwd)
DIR_BASE = Path(__file__).resolve().parent
RUTA_PROVEEDORES = DIR_BASE / "data" / "Proveedores.xlsx"
RUTA_TIENDAS = DIR_BASE / "data" / "Unidad_de_Negocio.xlsx"

# ------------------------------------------------------------
# MAPEO DE TIENDAS (codigo -> (unidad_negocio, centro_costo, subsidiaria))
# NOTA: esta informacion tambien existe en data/Unidad_de_Negocio.xlsx.
# Mantener ambas fuentes sincronizadas o eliminar una de las dos.
# ------------------------------------------------------------
MAPEO_TIENDAS = {
    # Guatemala (subsidiaria 15)
    601: (82, 241, 15),
    602: (85, 242, 15),
    603: (88, 243, 15),
    605: (91, 244, 15),
    606: (94, 245, 15),
    607: (97, 246, 15),
    608: (100, 247, 15),
    609: (103, 248, 15),
    610: (106, 249, 15),

    # Costa Rica (subsidiaria 18)
    620: (43, 229, 18),
    621: (46, 233, 18),
    622: (49, 231, 18),
    623: (52, 234, 18),
    624: (55, 230, 18),
    625: (58, 227, 18),
    626: (61, 232, 18),
    627: (64, 228, 18),

    # El Salvador (subsidiaria 16)
    641: (13, 149, 16),
    642: (14, 147, 16),
    643: (15, 148, 16),

    # Honduras (subsidiaria 17)
    651: (69, 235, 17),
    652: (72, 236, 17),
    653: (75, 237, 17),
    654: (78, 238, 17),

    # Panama (subsidiaria 19)
    671: (26, 218, 19),
    673: (29, 219, 19),
    674: (32, 220, 19),
    675: (35, 221, 19),
    690: (38, 223, 19),
}

SUBSIDIARIA_TEXTO = {
    15: "0211 OD GUATEMALA Y COMPAÑIA LIMITADA",
    16: "0213 OD EL SALVADOR LTDA, DE C.V.",
    17: "0216 OD HONDURAS S DE R L",
    18: "0214 OD ERIAL BQ S.A",
    19: "0217 OD PANAMA, S.A.",
}

NOMBRE_HOJA = {15: "OT-GT", 16: "OT-SV", 17: "OT-HN", 18: "OT-CR", 19: "OT-PA"}

TIENDAS_POR_HOJA = {
    "GT": [601, 602, 605, 607, 608, 610],
    "SV": [642],
    "HN": [651, 652],
    "CR": [620, 621, 622, 623, 625],
    "PA": [671, 675],
}

TIENDAS_POR_HOJA_MX = {
    "GT-MX": [601, 602, 605, 607, 608, 610],
    "SV-MX": [642],
    "HN-MX": [651, 652],
    "CR-MX": [620, 621, 622, 623, 625],
}

TIPOS_COMPRA = [
    "Compras Reabasto",
    "Inventario Primera Vez (Local e Importación)",
    "Inventariable Bajo Pedido (Local e Importación)",
    "Pie de Camión",
]

COLUMNAS_OT = [
    "ID_EXTERNO", "FECHA", "SUBSIDIARIA", "UNIDAD_ORIGEN", "UNIDAD_DESTINO",
    "EMPLEADO", "TRANSPORTISTA", "ID_INTERNAL", "SKU_NETSUIT", "CANTIDAD",
    "CENTRO_COSTO",
]

ENCABEZADOS_OT = [
    "ID EXTERNO", "FECHA", "SUBSIDIARIA", "UNIDAD DE NEGOCIO DE ORIGEN",
    "UNIDAD DE NEGOCIO DE DESTINO", "EMPLEADO", "TRANSPORTISTA", "ID INTERNAL",
    "SKU NETSUIT", "CANTIDAD", "CENTRO DE COSTO",
]

COLUMNAS_OC = [
    "EXTERNAL ID", "PROVEEDOR", "NOMBRE PROVEDOR", "FECHA",
    "TIPO DE COMPRA OD", "NOTA", "MONEDA", "UNIDAD DE NEGOCIO",
    "CENTRO DE COSTO", "SUBSIDIARIA", "ARTICULO", "CANTIDAD",
    "COSTO", "UNIDAD DE NEGOCIO_2", "CENTRO DE COSTO_2",
    "validador.tiendanombre", "validador.proveedornombre",
]

# NetSuite espera dos columnas repetidas; se emiten con el nombre duplicado.
ENCABEZADOS_OC_CSV = [
    "EXTERNAL ID", "PROVEEDOR", "NOMBRE PROVEDOR", "FECHA",
    "TIPO DE COMPRA OD", "NOTA", "MONEDA", "UNIDAD DE NEGOCIO",
    "CENTRO DE COSTO", "SUBSIDIARIA", "ARTICULO", "CANTIDAD",
    "COSTO", "UNIDAD DE NEGOCIO", "CENTRO DE COSTO",
    "validador.tiendanombre", "validador.proveedornombre",
]

SIN_MATCH = "#SIN_MATCH"


# ------------------------------------------------------------
# UTILIDADES
# ------------------------------------------------------------
def fecha_serial_excel(fecha_date):
    """Convierte una fecha a numero serial de Excel."""
    base = datetime(1899, 12, 30).date()
    return (fecha_date - base).days


def parsear_cantidad(valor):
    """
    Convierte texto a float manejando separadores de miles y decimales.

    Acepta formatos como: '1234', '1,234', '1.234,5', '1,234.5', '1 234'.
    Devuelve None si el valor no es interpretable como numero.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None

    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return None

    negativo = texto.startswith("-")
    texto = re.sub(r"[^\d,.]", "", texto)
    if not texto:
        return None

    ultima_coma = texto.rfind(",")
    ultimo_punto = texto.rfind(".")

    if ultima_coma >= 0 and ultimo_punto >= 0:
        # Ambos presentes: el ultimo en aparecer es el decimal
        if ultima_coma > ultimo_punto:
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif ultima_coma >= 0:
        # Solo comas. Ambiguo: '1,234' puede ser 1234 o 1.234.
        # Se asume separador de miles cuando hay exactamente 3 digitos
        # despues de la ultima coma (convencion de miles).
        decimales = len(texto) - ultima_coma - 1
        if decimales == 3 and texto.count(",") >= 1:
            texto = texto.replace(",", "")
        else:
            texto = texto.replace(",", ".")
    # Solo puntos o sin separadores: se deja tal cual

    try:
        numero = float(texto)
    except ValueError:
        return None

    return -numero if negativo else numero


def formatear_cantidad(valor, decimales=4):
    """
    Redondea y devuelve int si el valor es entero, si no float.

    El redondeo evita que el residuo de punto flotante acumulado durante
    el emparejamiento (ej. 0.09999999999999998) llegue al archivo final.
    """
    if valor is None:
        return valor
    valor = round(float(valor), decimales)
    if valor.is_integer():
        return int(valor)
    return valor


# ------------------------------------------------------------
# TRANSFERENCIAS
# ------------------------------------------------------------
def procesar_movimientos(df_mov, mapeo, debug_container, sheet_name=""):
    """
    Empareja cantidades negativas (origen) con positivas (destino) por SKU.

    Devuelve (transferencias, incidencias).
    """
    incidencias = []

    tienda_cols = []
    for col in df_mov.columns:
        primera_parte = str(col).split(" ", 1)[0]
        if primera_parte.isdigit():
            tienda_cols.append(col)

    debug_container.write(f"📊 Columnas detectadas como tiendas: {tienda_cols}")

    if not tienda_cols:
        debug_container.error(
            "❌ No se detectaron columnas numéricas. Los encabezados deben "
            "comenzar con un número (ej. '601', '602 TIENDA X')."
        )
        return [], ["Sin columnas de tienda detectadas"]

    columnas_requeridas = {"SKU", "ID interno"}
    faltantes = columnas_requeridas - set(df_mov.columns)
    if faltantes:
        debug_container.error(f"❌ Faltan columnas obligatorias: {sorted(faltantes)}")
        return [], [f"Columnas faltantes: {sorted(faltantes)}"]

    transfers = []
    total_filas = 0

    for _, row in df_mov.iterrows():
        total_filas += 1
        sku = row["SKU"]
        id_interno = row["ID interno"]
        origenes = []
        destinos = []

        for col in tienda_cols:
            cantidad = parsear_cantidad(row[col])
            if cantidad is None or abs(cantidad) < TOLERANCIA:
                continue
            id_tienda = int(str(col).split(" ")[0])
            if cantidad < 0:
                origenes.append([id_tienda, -cantidad])
            else:
                destinos.append([id_tienda, cantidad])

        if not origenes and not destinos:
            continue

        total_origen = sum(c for _, c in origenes)
        total_destino = sum(c for _, c in destinos)
        if abs(total_origen - total_destino) > 0.001:
            msg = f"SKU {sku}: desbalanceado (origen {total_origen} vs destino {total_destino})"
            debug_container.warning(f"⚠️ {msg}")
            incidencias.append(msg)
            continue

        orig = [list(o) for o in origenes]
        dest = [list(d) for d in destinos]

        while orig and dest:
            orig.sort(key=lambda x: x[1], reverse=True)
            dest.sort(key=lambda x: x[1], reverse=True)
            o_id, o_cant = orig[0]
            d_id, d_cant = dest[0]

            if o_id not in mapeo:
                msg = f"Tienda origen {o_id} no está en mapeo (SKU {sku})"
                debug_container.error(f"❌ {msg}")
                incidencias.append(msg)
                break
            if d_id not in mapeo:
                msg = f"Tienda destino {d_id} no está en mapeo (SKU {sku})"
                debug_container.error(f"❌ {msg}")
                incidencias.append(msg)
                break

            unidad_origen, centro_origen, sub_origen = mapeo[o_id]
            unidad_destino, _, sub_destino = mapeo[d_id]

            # FIX: antes usaba 'continue', lo que reevaluaba el mismo par
            # indefinidamente (bucle infinito). Se aborta el SKU completo.
            if sub_origen != sub_destino:
                msg = (
                    f"Transferencia entre países: {o_id}(sub{sub_origen}) → "
                    f"{d_id}(sub{sub_destino}) - SKU {sku} omitida"
                )
                debug_container.error(f"🌎 {msg}")
                incidencias.append(msg)
                break

            transferir = min(o_cant, d_cant)

            transfers.append({
                "origen_hoja": sheet_name,
                "subsidiaria_num": sub_origen,
                "UNIDAD_ORIGEN": unidad_origen,
                "UNIDAD_DESTINO": unidad_destino,
                "CENTRO_COSTO": centro_origen,
                "ID_INTERNAL": id_interno,
                "SKU_NETSUIT": sku,
                "CANTIDAD": formatear_cantidad(transferir),
            })

            # FIX: comparacion con tolerancia en vez de '== 0'. Con decimales,
            # el residuo de punto flotante impedia el pop y colgaba el bucle.
            orig[0][1] -= transferir
            if orig[0][1] < TOLERANCIA:
                orig.pop(0)
            dest[0][1] -= transferir
            if dest[0][1] < TOLERANCIA:
                dest.pop(0)

    debug_container.info(
        f"✅ Filas procesadas: {total_filas}. Transferencias generadas: {len(transfers)}"
    )
    return transfers, incidencias


def generar_excel_bytes(transfers):
    """Genera el libro de Excel con una hoja por hoja de origen."""
    if not transfers:
        return None

    grupos = {}
    for t in transfers:
        grupos.setdefault(t["origen_hoja"], []).append(t)

    wb = Workbook()
    wb.remove(wb.active)
    fecha_serial = fecha_serial_excel(datetime.now().date())

    for hoja_nombre, lista in grupos.items():
        df = pd.DataFrame(lista)

        # FIX: el ID externo anterior era OT{fecha}{origen}{destino}, identico
        # para todos los SKU del mismo par de tiendas en el mismo dia. Se
        # agregan separadores y un correlativo para garantizar unicidad.
        df["ID_EXTERNO"] = [
            "OT-{}-{}-{}-{:04d}".format(
                fecha_serial, fila["UNIDAD_ORIGEN"], fila["UNIDAD_DESTINO"], i + 1
            )
            for i, fila in enumerate(lista)
        ]
        df["FECHA"] = fecha_serial

        sub_num = lista[0]["subsidiaria_num"]
        df["SUBSIDIARIA"] = SUBSIDIARIA_TEXTO.get(sub_num, f"SUBSIDIARIA_{sub_num}")
        df["EMPLEADO"] = ""
        df["TRANSPORTISTA"] = "TRANSPORTE PROPIO"

        df = df[COLUMNAS_OT]
        df = df.sort_values(by="ID_EXTERNO", ascending=False).reset_index(drop=True)

        ws = wb.create_sheet(title=str(hoja_nombre)[:31])
        ws.append(ENCABEZADOS_OT)
        for _, row in df.iterrows():
            ws.append(list(row.values))

        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ------------------------------------------------------------
# CATALOGOS
# ------------------------------------------------------------
@st.cache_data
def cargar_proveedores():
    """
    Carga el catalogo de proveedores.

    Devuelve (lookup, error). No emite UI: al estar cacheada, un st.error()
    dentro de esta funcion no volveria a mostrarse en llamadas posteriores.
    """
    try:
        df = pd.read_excel(RUTA_PROVEEDORES, dtype=str)
        lookup = dict(
            zip(df["Proveedor"].str.strip(), df["ID_PROVEEDOR"].str.strip())
        )
        return lookup, None
    except Exception as e:
        return {}, f"No se pudo cargar {RUTA_PROVEEDORES.name}: {e}"


@st.cache_data
def cargar_tiendas():
    """Carga el catalogo de tiendas. Devuelve (lookup, error)."""
    try:
        df = pd.read_excel(RUTA_TIENDAS, dtype=str)
        lookup = {}
        for _, row in df.iterrows():
            num = str(row["No. Tienda"]).strip()
            lookup[num] = {
                "unidad_negocio": str(row.get("UNIDAD DE NEGOCIO", "")).strip(),
                "centro_costo": str(row.get("CENTRO DE COSTO", "")).strip(),
                "subsidiaria": str(row.get("SUBSIDARIA", "")).strip(),
                "nombre_tienda": str(row.get("Unidad de Negocio del inventario", "")).strip(),
            }
        return lookup, None
    except Exception as e:
        return {}, f"No se pudo cargar {RUTA_TIENDAS.name}: {e}"


# ------------------------------------------------------------
# ORDENES DE COMPRA
# ------------------------------------------------------------
def procesar_archivo_pedido(uploaded_file, tipo_compra, tiendas_por_hoja):
    """
    Procesa el archivo de pedido y devuelve (df, avisos, proveedores_sin_match).
    """
    avisos = []
    sin_match = set()
    hoy = datetime.now().strftime("%d/%m/%Y")
    hoy_compacto = hoy.replace("/", "")
    filas = []

    lookup_prov, error_prov = cargar_proveedores()
    lookup_tienda, error_tienda = cargar_tiendas()
    if error_prov:
        avisos.append(error_prov)
    if error_tienda:
        avisos.append(error_tienda)

    correlativo = 0

    for hoja, tiendas in tiendas_por_hoja.items():
        try:
            df_raw = pd.read_excel(
                uploaded_file, sheet_name=hoja, header=None, dtype=str
            )
        except Exception as e:
            avisos.append(f"No se pudo leer la hoja '{hoja}': {e}")
            continue

        if len(df_raw) <= FILA_PRIMER_DATO:
            avisos.append(
                f"La hoja '{hoja}' tiene menos de {FILA_PRIMER_DATO + 1} filas, se omite."
            )
            continue

        header_row = df_raw.iloc[FILA_ENCABEZADOS]
        col_index_map = {}
        for idx, val in enumerate(header_row):
            if pd.isna(val):
                continue
            val_str = str(val).strip()
            if val_str.isdigit():
                col_index_map[int(val_str)] = idx

        tiendas_faltantes = [t for t in tiendas if t not in col_index_map]
        if tiendas_faltantes:
            avisos.append(
                f"En la hoja '{hoja}' faltan las tiendas: {tiendas_faltantes}. Se omitirán."
            )
        tiendas_a_usar = [t for t in tiendas if t in col_index_map]

        for idx_fila in range(FILA_PRIMER_DATO, len(df_raw)):
            fila = df_raw.iloc[idx_fila]

            id_interno = str(fila[COL_ID_INTERNO]).strip() if pd.notna(fila[COL_ID_INTERNO]) else ""
            if not id_interno or id_interno.lower() == "nan":
                continue

            nombre_proveedor = (
                str(fila[COL_PROVEEDOR]).strip() if pd.notna(fila[COL_PROVEEDOR]) else ""
            )
            id_proveedor = lookup_prov.get(nombre_proveedor, SIN_MATCH)
            if id_proveedor == SIN_MATCH and nombre_proveedor:
                sin_match.add(nombre_proveedor)

            for tienda in tiendas_a_usar:
                cantidad = parsear_cantidad(fila[col_index_map[tienda]])
                if cantidad is None or cantidad <= 0:
                    continue
                cantidad = formatear_cantidad(cantidad)

                datos_tienda = lookup_tienda.get(str(tienda), {})
                unidad_neg = datos_tienda.get("unidad_negocio", "")
                centro_costo = datos_tienda.get("centro_costo", "")
                subsidiaria = datos_tienda.get("subsidiaria", "")
                nombre_tienda = datos_tienda.get("nombre_tienda", "")

                # FIX: el ID anterior era OcBrian{prov}{tienda}{fecha}, identico
                # para todos los articulos del mismo proveedor/tienda/dia y
                # ambiguo por concatenacion sin separador. Se agregan guiones
                # y un correlativo global.
                correlativo += 1
                external_id = (
                    f"OC-{id_proveedor}-{tienda}-{hoy_compacto}-{correlativo:05d}"
                )

                filas.append({
                    "EXTERNAL ID": external_id,
                    "PROVEEDOR": id_proveedor,
                    "NOMBRE PROVEDOR": "",
                    "FECHA": hoy,
                    "TIPO DE COMPRA OD": tipo_compra,
                    "NOTA": "",
                    "MONEDA": "US Dollar",
                    "UNIDAD DE NEGOCIO": unidad_neg,
                    "CENTRO DE COSTO": centro_costo,
                    "SUBSIDIARIA": subsidiaria,
                    "ARTICULO": id_interno,
                    "CANTIDAD": cantidad,
                    "COSTO": "",
                    "UNIDAD DE NEGOCIO_2": unidad_neg,
                    "CENTRO DE COSTO_2": centro_costo,
                    "validador.tiendanombre": nombre_tienda,
                    "validador.proveedornombre": nombre_proveedor,
                })

    if not filas:
        return pd.DataFrame(), avisos, sin_match

    df_final = pd.DataFrame(filas)[COLUMNAS_OC]
    df_final = df_final.sort_values(by="EXTERNAL ID").reset_index(drop=True)
    return df_final, avisos, sin_match


def generar_csv(df, separador):
    """
    Serializa el DataFrame a CSV usando el modulo csv.

    FIX: el escapado manual anterior no manejaba comillas dentro de los
    valores y no citaba nada cuando el separador era ';'.
    """
    buffer = io.StringIO()
    writer = csv.writer(buffer, delimiter=separador, quoting=csv.QUOTE_MINIMAL,
                        lineterminator="\n")
    writer.writerow(ENCABEZADOS_OC_CSV)
    for _, row in df.iterrows():
        writer.writerow([row[c] for c in COLUMNAS_OC])
    return buffer.getvalue()


# ------------------------------------------------------------
# PAGINAS
# ------------------------------------------------------------
def pagina_transferencias():
    st.title("📦 Generador de OT desde movimientos")

    uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])
    if not uploaded_file:
        st.info("Sube un archivo Excel para comenzar.")
        return

    try:
        todas_hojas = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        return

    for hoja in todas_hojas:
        todas_hojas[hoja].columns = [str(c) for c in todas_hojas[hoja].columns]

    st.success(f"✅ Archivo cargado con {len(todas_hojas)} hoja(s): {list(todas_hojas.keys())}")

    primera_hoja = list(todas_hojas.keys())[0]
    st.subheader(f"Vista previa de la hoja '{primera_hoja}' (primeras 10 filas)")
    st.write(todas_hojas[primera_hoja].head(10))

    debug_container = st.container()
    with debug_container:
        st.subheader("📝 Log de procesamiento")

    if not st.button("🚀 Procesar"):
        return

    with st.spinner("Procesando todas las hojas..."):
        all_transfers = []
        todas_incidencias = []
        for sheet_name, df_mov in todas_hojas.items():
            st.write(f"📄 Procesando hoja: **{sheet_name}**")
            transfers_hoja, incidencias = procesar_movimientos(
                df_mov, MAPEO_TIENDAS, debug_container, sheet_name=sheet_name
            )
            st.write(f"   ↳ {len(transfers_hoja)} transferencias generadas")
            all_transfers.extend(transfers_hoja)
            todas_incidencias.extend(incidencias)

    if not all_transfers:
        st.error("No se generaron transferencias. Revisa el log de arriba.")
        return

    if todas_incidencias:
        st.warning(f"⚠️ Se registraron {len(todas_incidencias)} incidencia(s). "
                   "Revisa el log antes de importar a NetSuite.")

    excel_data = generar_excel_bytes(all_transfers)
    st.success(f"✅ Se generaron {len(all_transfers)} transferencias en total")
    st.download_button(
        "📥 Descargar Excel",
        excel_data,
        f"OT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def pagina_pedidos(titulo, tiendas_por_hoja, sufijo_archivo="", key_prefix="oc"):
    """
    Pagina de ordenes de compra.

    FIX: antes existian dos funciones casi identicas (pagina_pedidos y
    pagina_pedidos_mx) con ~90 lineas duplicadas. Ahora es una sola,
    parametrizada por mapeo de tiendas y prefijo de claves de widget.
    """
    st.title(titulo)
    st.markdown(
        "Sube el archivo `Nuevo Análisis V2.xlsx` para generar el archivo de "
        f"órdenes de compra. Hojas esperadas: {list(tiendas_por_hoja.keys())}"
    )

    separador = st.selectbox(
        "Separador del CSV",
        options=[",", ";"],
        format_func=lambda x: 'Coma  ","' if x == "," else 'Punto y coma  ";"',
        index=0,
        key=f"sep_{key_prefix}",
    )

    tipo_compra = st.selectbox(
        "Tipo de compra", options=TIPOS_COMPRA, index=0, key=f"tipo_{key_prefix}"
    )

    uploaded_file = st.file_uploader(
        "Cargar archivo Excel", type=["xlsx"], key=f"upload_{key_prefix}"
    )

    if uploaded_file is None:
        st.info("Por favor, sube un archivo Excel para comenzar.")
        return

    with st.spinner("Procesando archivo..."):
        df_resultado, avisos, sin_match = procesar_archivo_pedido(
            uploaded_file, tipo_compra, tiendas_por_hoja
        )

    for aviso in avisos:
        st.warning(f"⚠️ {aviso}")

    if df_resultado.empty:
        st.error("No se pudo generar la orden de compra. Revisa el formato del archivo.")
        return

    # FIX: antes los proveedores sin coincidencia pasaban al archivo final
    # sin ninguna advertencia visible.
    if sin_match:
        st.error(
            f"❌ {len(sin_match)} proveedor(es) sin coincidencia en el catálogo. "
            f"Aparecerán como `{SIN_MATCH}` en el archivo y serán rechazados "
            "por NetSuite:"
        )
        st.write(sorted(sin_match))

    st.success(f"✅ Se generaron {len(df_resultado)} filas para la orden de compra.")
    st.subheader("📊 Vista previa")
    st.dataframe(df_resultado, use_container_width=True)

    csv_data = generar_csv(df_resultado, separador)
    st.download_button(
        label="📥 Descargar CSV",
        data=csv_data,
        file_name=f"orden_compra{sufijo_archivo}_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True,
        key=f"descarga_{key_prefix}",
    )


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    st.set_page_config(page_title="Generador de Documentos", layout="wide")

    st.sidebar.title("Navegación")
    opcion = st.sidebar.radio(
        "¿Qué deseas hacer?",
        ("Formatear Transferencia", "Formatear Orden de Pedido", "Formatear Orden de Pedido (MX)"),
    )

    if opcion == "Formatear Transferencia":
        pagina_transferencias()
    elif opcion == "Formatear Orden de Pedido":
        pagina_pedidos(
            "📦 Generador de Orden de Compra (estándar)",
            TIENDAS_POR_HOJA,
            sufijo_archivo="",
            key_prefix="oc",
        )
    else:
        pagina_pedidos(
            "📦 Generador de Orden de Compra (MX)",
            TIENDAS_POR_HOJA_MX,
            sufijo_archivo="_MX",
            key_prefix="oc_mx",
        )


if __name__ == "__main__":
    main()
