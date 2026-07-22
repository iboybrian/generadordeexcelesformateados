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

# Columnas extra que se agregan solo si se cargo un archivo de stock
COLUMNAS_OT_STOCK = [
    "STOCK_INICIAL_ORIGEN", "STOCK_NUEVO_ORIGEN",
    "STOCK_INICIAL_DESTINO", "STOCK_NUEVO_DESTINO",
]

ENCABEZADOS_OT = [
    "ID EXTERNO", "FECHA", "SUBSIDIARIA", "UNIDAD DE NEGOCIO DE ORIGEN",
    "UNIDAD DE NEGOCIO DE DESTINO", "EMPLEADO", "TRANSPORTISTA", "ID INTERNAL",
    "SKU NETSUIT", "CANTIDAD", "CENTRO DE COSTO",
]

ENCABEZADOS_OT_STOCK = [
    "STOCK INICIAL ORIGEN", "STOCK NUEVO ORIGEN",
    "STOCK INICIAL DESTINO", "STOCK NUEVO DESTINO",
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
# ARCHIVO DE STOCK (CSV opcional para transferencias)
# ------------------------------------------------------------
# Nombres de columna aceptados. Se busca la primera que exista, ya sea con
# acentos correctos o con mojibake (UTF-8 leido como Latin-1).
COL_STOCK_SKU = ["SKU"]
COL_STOCK_UBICACION = ["Ubicación del inventario", "UbicaciÃ³n del inventario"]
COL_STOCK_FISICO = ["Físico en ubicación", "FÃ­sico en ubicaciÃ³n"]


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


def reparar_mojibake(texto):
    """
    Repara texto UTF-8 que fue leido como Latin-1 ('PRÃ“CERES' -> 'PRÓCERES').

    Si el texto ya es correcto, la operacion falla y se devuelve intacto,
    por lo que es seguro aplicarla siempre.
    """
    if not isinstance(texto, str):
        return texto
    try:
        return texto.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return texto


def resolver_columna(df, nombres_posibles):
    """Devuelve el primer nombre de columna que exista en el DataFrame."""
    for nombre in nombres_posibles:
        if nombre in df.columns:
            return nombre
    return None


def extraer_numero_tienda(texto):
    """
    Extrae el numero de tienda de 'OD | GT | 601 TIENDA MAJADAS' -> 601.

    Toma el primer grupo de digitos que aparezca despues del ultimo '|'.
    Devuelve None si no encuentra ninguno.
    """
    if not isinstance(texto, str):
        return None
    parte = texto.split("|")[-1].strip()
    match = re.search(r"\d+", parte)
    return int(match.group()) if match else None


def cargar_stock(archivo_csv, skus_relevantes=None):
    """
    Carga el CSV de stock y devuelve (stock, avisos).

    stock: dict {(sku, num_tienda): cantidad_float}
    Solo incluye tiendas presentes en MAPEO_TIENDAS; el resto se descarta
    (bodegas externas como 20601, tiendas de otros paises, etc).

    Una celda de stock vacia se interpreta como 0, no como ausente: la
    tienda aparece en el archivo, por lo tanto su existencia es cero.

    skus_relevantes: si se proporciona un set de SKU, se descartan las filas
    de cualquier otro SKU antes de procesar. Reduce memoria cuando el
    catalogo es mucho mayor que el archivo de movimientos.

    Implementado con operaciones vectorizadas de pandas: recorrer el
    DataFrame con iterrows() tardaba ~40s en un archivo de 300k filas.
    """
    avisos = []

    try:
        df = pd.read_csv(archivo_csv, dtype=str, keep_default_na=False)
    except Exception as e:
        return {}, [f"No se pudo leer el CSV de stock: {e}"]

    df.columns = [reparar_mojibake(str(c)) for c in df.columns]

    col_sku = resolver_columna(df, COL_STOCK_SKU)
    col_ubi = resolver_columna(df, [reparar_mojibake(c) for c in COL_STOCK_UBICACION])
    col_fis = resolver_columna(df, [reparar_mojibake(c) for c in COL_STOCK_FISICO])

    faltantes = []
    if col_sku is None:
        faltantes.append("SKU")
    if col_ubi is None:
        faltantes.append("Ubicación del inventario")
    if col_fis is None:
        faltantes.append("Físico en ubicación")
    if faltantes:
        return {}, [
            f"El CSV de stock no tiene las columnas requeridas: {faltantes}. "
            f"Columnas encontradas: {list(df.columns)[:10]}..."
        ]

    # Solo las tres columnas que interesan
    df = df[[col_sku, col_ubi, col_fis]].copy()
    df.columns = ["sku", "ubicacion", "fisico"]

    df["sku"] = df["sku"].astype(str).str.strip()
    df = df[(df["sku"] != "") & (df["sku"].str.lower() != "nan")]

    # Filtro opcional por SKU presentes en el archivo de movimientos
    if skus_relevantes:
        antes = len(df)
        df = df[df["sku"].isin(skus_relevantes)]
        avisos.append(
            f"Filtrado por SKU del archivo de movimientos: {antes:,} → {len(df):,} filas."
        )

    if df.empty:
        return {}, avisos + ["El CSV de stock no tiene filas utilizables."]

    # Numero de tienda: digitos despues del ultimo '|', vectorizado
    ubic = df["ubicacion"].astype(str).map(reparar_mojibake)
    df["num_tienda"] = pd.to_numeric(
        ubic.str.rsplit("|", n=1).str[-1].str.extract(r"(\d+)", expand=False),
        errors="coerce",
    )

    validas = df["num_tienda"].isin(MAPEO_TIENDAS.keys())
    descartadas = sorted(
        int(x) for x in df.loc[~validas & df["num_tienda"].notna(), "num_tienda"].unique()
    )
    df = df[validas].copy()
    df["num_tienda"] = df["num_tienda"].astype(int)

    if df.empty:
        if descartadas:
            avisos.append(
                f"Se ignoraron ubicaciones fuera del mapeo de tiendas: {descartadas}"
            )
        return {}, avisos + ["Ninguna ubicación del CSV coincide con el mapeo de tiendas."]

    # Regla acordada: si SKU+tienda se repite, conservar el primero
    antes_dedup = len(df)
    df = df.drop_duplicates(subset=["sku", "num_tienda"], keep="first")
    duplicadas = antes_dedup - len(df)

    # Cantidad: los datos son enteros sin separadores, asi que to_numeric
    # directo es suficiente; lo no numerico o vacio queda en 0.
    cantidades = pd.to_numeric(
        df["fisico"].astype(str).str.strip().replace("", "0"), errors="coerce"
    ).fillna(0.0)

    stock = dict(zip(zip(df["sku"], df["num_tienda"]), cantidades.astype(float)))

    if descartadas:
        avisos.append(
            f"Se ignoraron ubicaciones fuera del mapeo de tiendas: {descartadas}"
        )
    if duplicadas:
        avisos.append(
            f"Se encontraron {duplicadas} combinación(es) SKU+tienda repetidas; "
            "se conservó la primera."
        )

    return stock, avisos


def calcular_stock_resultante(transfers, stock):
    """
    Agrega columnas de stock inicial y resultante a cada transferencia.

    Aplica las salidas y entradas de forma acumulativa en el orden en que
    aparecen las transferencias, de modo que una tienda que envia varias
    veces refleje el saldo corriente.

    Devuelve la lista de incidencias (saldos negativos).
    """
    incidencias = []
    saldo = dict(stock)          # copia mutable, se va descontando
    conocidos = set(stock)       # claves presentes en el archivo original

    for t in transfers:
        sku = str(t["SKU_NETSUIT"]).strip()
        clave_origen = (sku, t["TIENDA_ORIGEN"])
        clave_destino = (sku, t["TIENDA_DESTINO"])
        cantidad = float(t["CANTIDAD"])

        # ORIGEN
        if clave_origen in conocidos:
            inicial = saldo.get(clave_origen, 0.0)
            resultante = inicial - cantidad
            saldo[clave_origen] = resultante
            t["STOCK_INICIAL_ORIGEN"] = formatear_cantidad(inicial)
            t["STOCK_NUEVO_ORIGEN"] = formatear_cantidad(resultante)
            if resultante < 0:
                incidencias.append(
                    f"SKU {sku}: transferir {formatear_cantidad(cantidad)} de la tienda "
                    f"{t['TIENDA_ORIGEN']} a la {t['TIENDA_DESTINO']} deja saldo "
                    f"negativo ({formatear_cantidad(resultante)})"
                )
        else:
            # SKU+tienda ausente del archivo de stock: se deja en blanco
            t["STOCK_INICIAL_ORIGEN"] = ""
            t["STOCK_NUEVO_ORIGEN"] = ""

        # DESTINO
        if clave_destino in conocidos:
            inicial = saldo.get(clave_destino, 0.0)
            resultante = inicial + cantidad
            saldo[clave_destino] = resultante
            t["STOCK_INICIAL_DESTINO"] = formatear_cantidad(inicial)
            t["STOCK_NUEVO_DESTINO"] = formatear_cantidad(resultante)
        else:
            t["STOCK_INICIAL_DESTINO"] = ""
            t["STOCK_NUEVO_DESTINO"] = ""

    return incidencias


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

    # to_dict("records") es mucho mas rapido que iterrows() en archivos grandes
    for row in df_mov.to_dict("records"):
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
                "TIENDA_ORIGEN": o_id,
                "TIENDA_DESTINO": d_id,
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


def generar_excel_bytes(transfers, con_stock=False):
    """
    Genera el libro de Excel con una hoja por hoja de origen.

    El ID externo es COMPARTIDO por par origen-destino: todas las lineas de
    un mismo par llevan el mismo codigo, porque NetSuite agrupa por ese ID
    para armar un solo documento de transferencia. Las filas se ordenan por
    ID externo para que las lineas del mismo documento queden contiguas.

    Si con_stock es True, agrega las columnas de stock inicial/resultante.
    Los saldos ya vienen calculados en el orden real de procesamiento; el
    orden estable conserva su secuencia dentro de cada grupo.
    """
    if not transfers:
        return None

    columnas = COLUMNAS_OT + (COLUMNAS_OT_STOCK if con_stock else [])
    encabezados = ENCABEZADOS_OT + (ENCABEZADOS_OT_STOCK if con_stock else [])

    grupos = {}
    for t in transfers:
        grupos.setdefault(t["origen_hoja"], []).append(t)

    wb = Workbook()
    wb.remove(wb.active)
    fecha_serial = fecha_serial_excel(datetime.now().date())

    for hoja_nombre, lista in grupos.items():
        df = pd.DataFrame(lista)

        # ID externo agrupador: identico para todas las lineas del mismo par
        # origen-destino. NetSuite lo usa como clave de agrupacion, por lo que
        # NO debe ser unico por fila.
        df["ID_EXTERNO"] = [
            f"OT{fecha_serial}{fila['UNIDAD_ORIGEN']}{fila['UNIDAD_DESTINO']}"
            for fila in lista
        ]
        df["FECHA"] = fecha_serial

        sub_num = lista[0]["subsidiaria_num"]
        df["SUBSIDIARIA"] = SUBSIDIARIA_TEXTO.get(sub_num, f"SUBSIDIARIA_{sub_num}")
        df["EMPLEADO"] = ""
        df["TRANSPORTISTA"] = "TRANSPORTE PROPIO"

        df = df[columnas]
        # kind="stable": agrupa por ID sin alterar el orden relativo dentro
        # de cada grupo, para que la secuencia de saldos siga siendo legible.
        df = df.sort_values(
            by="ID_EXTERNO", ascending=True, kind="stable"
        ).reset_index(drop=True)

        ws = wb.create_sheet(title=str(hoja_nombre)[:31])
        ws.append(encabezados)
        for fila in df.itertuples(index=False, name=None):
            ws.append(list(fila))

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
        for row in df.to_dict("records"):
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

        # .iloc por fila es lento; se materializa el bloque de datos una vez
        bloque = df_raw.iloc[FILA_PRIMER_DATO:].to_numpy()

        for fila in bloque:

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

                # EXTERNAL ID agrupador: identico para todas las lineas del
                # mismo proveedor + tienda + fecha. NetSuite lo usa como clave
                # de agrupacion para armar una sola orden de compra, por lo
                # que NO debe ser unico por fila.
                external_id = f"OcBrian{id_proveedor}{tienda}{hoy_compacto}"

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
    # kind="stable": agrupa las lineas del mismo EXTERNAL ID de forma contigua
    # sin alterar el orden en que fueron leidas dentro de cada grupo.
    df_final = df_final.sort_values(
        by="EXTERNAL ID", kind="stable"
    ).reset_index(drop=True)
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
    for fila in df[COLUMNAS_OC].itertuples(index=False, name=None):
        writer.writerow(fila)
    return buffer.getvalue()


# ------------------------------------------------------------
# PAGINAS
# ------------------------------------------------------------
def pagina_transferencias():
    st.title("📦 Generador de OT desde movimientos")

    uploaded_file = st.file_uploader("Sube tu archivo Excel de movimientos", type=["xlsx"])

    st.markdown("---")
    st.subheader("📊 Stock (opcional)")
    st.caption(
        "Sube el CSV de existencias para calcular el saldo resultante en la "
        "tienda de origen y destino. Requiere las columnas `SKU`, "
        "`Ubicación del inventario` y `Físico en ubicación`."
    )
    stock_file = st.file_uploader("Sube el CSV de stock", type=["csv"], key="stock_csv")
    st.markdown("---")

    if not uploaded_file:
        st.info("Sube un archivo Excel de movimientos para comenzar.")
        return

    try:
        todas_hojas = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        return

    for hoja in todas_hojas:
        todas_hojas[hoja].columns = [str(c) for c in todas_hojas[hoja].columns]

    st.success(f"✅ Archivo cargado con {len(todas_hojas)} hoja(s): {list(todas_hojas.keys())}")

    # SKU presentes en el archivo de movimientos: sirven para descartar de
    # entrada las filas irrelevantes del catalogo de stock.
    skus_movimientos = set()
    for df_hoja in todas_hojas.values():
        if "SKU" in df_hoja.columns:
            skus_movimientos.update(
                df_hoja["SKU"].dropna().astype(str).str.strip().tolist()
            )
    skus_movimientos.discard("")

    stock = {}
    if stock_file is not None:
        with st.spinner("Cargando stock..."):
            stock, avisos_stock = cargar_stock(
                stock_file, skus_relevantes=skus_movimientos or None
            )
        for aviso in avisos_stock:
            st.warning(f"⚠️ {aviso}")
        if stock:
            st.success(f"✅ Stock cargado: {len(stock)} combinaciones SKU + tienda.")
        else:
            st.error("❌ No se cargó ningún registro de stock. Se omitirá el cálculo.")

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

    incidencias_stock = []
    if stock:
        incidencias_stock = calcular_stock_resultante(all_transfers, stock)

    if todas_incidencias:
        st.warning(
            f"⚠️ Se registraron {len(todas_incidencias)} incidencia(s) de procesamiento. "
            "Revisa el log antes de importar a NetSuite."
        )

    if incidencias_stock:
        st.error(f"🔴 {len(incidencias_stock)} transferencia(s) dejan saldo negativo:")
        for msg in incidencias_stock:
            st.write(f"   • {msg}")
        st.caption(
            "Las transferencias se generan de todos modos; el saldo negativo "
            "queda reflejado en la columna STOCK NUEVO ORIGEN."
        )

    excel_data = generar_excel_bytes(all_transfers, con_stock=bool(stock))
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
