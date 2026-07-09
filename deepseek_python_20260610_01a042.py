import streamlit as st
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import io

# ------------------------------------------------------------
# MAPEO DE TIENDAS (código del encabezado → (unidad_negocio, centro_costo, subsidiaria_num))
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

    # Panamá (subsidiaria 19)
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

def fecha_serial_excel(fecha_date):
    base = datetime(1899, 12, 30).date()
    return (fecha_date - base).days

def procesar_movimientos(df_mov, mapeo, debug_container):
    """Procesa y muestra mensajes de depuración en el contenedor"""
    # Identificar columnas de tienda
    tienda_cols = []
    for col in df_mov.columns:
        try:
            parts = str(col).split(' ', 1)
            if parts[0].isdigit():
                tienda_cols.append(col)
        except:
            pass

    debug_container.write(f"📊 Columnas detectadas como tiendas: {tienda_cols}")

    if not tienda_cols:
        debug_container.error(
            "❌ No se detectaron columnas numéricas. Los encabezados deben comenzar con un número (ej. '601', '602 TIENDA X').")
        return []

    transfers = []
    total_filas = 0
    for idx, row in df_mov.iterrows():
        total_filas += 1
        sku = row['SKU']
        id_interno = row['ID interno']
        origenes = []
        destinos = []

        for col in tienda_cols:
            cantidad_raw = row[col]
            try:
                cantidad = float(cantidad_raw)
            except:
                cantidad = 0
            if pd.isna(cantidad) or cantidad == 0:
                continue
            id_tienda = int(str(col).split(' ')[0])
            if cantidad < 0:
                origenes.append([id_tienda, -cantidad])
            else:
                destinos.append([id_tienda, cantidad])

        if not origenes and not destinos:
            continue

        total_origen = sum(c for _, c in origenes)
        total_destino = sum(c for _, c in destinos)
        if abs(total_origen - total_destino) > 0.001:
            debug_container.warning(
                f"⚠️ SKU {sku}: desbalanceado (origen {total_origen} vs destino {total_destino})")
            continue

        # Emparejamiento
        orig = origenes.copy()
        dest = destinos.copy()
        while orig and dest:
            orig.sort(key=lambda x: x[1], reverse=True)
            dest.sort(key=lambda x: x[1], reverse=True)
            o_id, o_cant = orig[0]
            d_id, d_cant = dest[0]
            transferir = min(o_cant, d_cant)

            if o_id not in mapeo:
                debug_container.error(
                    f"❌ Tienda origen {o_id} no está en mapeo (SKU {sku})")
                break
            if d_id not in mapeo:
                debug_container.error(
                    f"❌ Tienda destino {d_id} no está en mapeo (SKU {sku})")
                break

            unidad_origen, centro_origen, sub_origen = mapeo[o_id]
            unidad_destino, _, sub_destino = mapeo[d_id]

            if sub_origen != sub_destino:
                debug_container.error(
                    f"🌎 Transferencia entre países: {o_id}(sub{sub_origen}) → {d_id}(sub{sub_destino}) - SKU {sku} omitida")
                continue

            transfers.append({
                'subsidiaria_num': sub_origen,
                'UNIDAD_ORIGEN': unidad_origen,
                'UNIDAD_DESTINO': unidad_destino,
                'CENTRO_COSTO': centro_origen,
                'ID_INTERNAL': id_interno,
                'SKU_NETSUIT': sku,
                'CANTIDAD': transferir,
            })

            orig[0][1] -= transferir
            if orig[0][1] == 0:
                orig.pop(0)
            dest[0][1] -= transferir
            if dest[0][1] == 0:
                dest.pop(0)

    debug_container.info(
        f"✅ Filas procesadas: {total_filas}. Transferencias generadas: {len(transfers)}")
    return transfers

def generar_excel_bytes(transfers):
    if not transfers:
        return None
    grupos = {}
    for t in transfers:
        sub = t['subsidiaria_num']
        grupos.setdefault(sub, []).append(t)
    wb = Workbook()
    wb.remove(wb.active)
    hoy = datetime.now().date()
    fecha_serial = fecha_serial_excel(hoy)
    for sub_num, lista in grupos.items():
        df = pd.DataFrame(lista)
        df['ID_EXTERNO'] = df.apply(
            lambda row: f"OT{fecha_serial}{row['UNIDAD_ORIGEN']}{row['UNIDAD_DESTINO']}", axis=1)
        df['FECHA'] = fecha_serial
        df['SUBSIDIARIA'] = SUBSIDIARIA_TEXTO.get(
            sub_num, f"SUBSIDIARIA_{sub_num}")
        df['EMPLEADO'] = ''
        df['TRANSPORTISTA'] = 'TRANSPORTE PROPIO'
        columnas_orden = [
            'ID_EXTERNO', 'FECHA', 'SUBSIDIARIA', 'UNIDAD_ORIGEN', 'UNIDAD_DESTINO',
            'EMPLEADO', 'TRANSPORTISTA', 'ID_INTERNAL', 'SKU_NETSUIT', 'CANTIDAD', 'CENTRO_COSTO'
        ]
        df = df[columnas_orden]
        sheet_name = NOMBRE_HOJA.get(sub_num, f"OT-{sub_num}")
        ws = wb.create_sheet(title=sheet_name)
        headers = [
            'ID EXTERNO', 'FECHA', 'SUBSIDIARIA', 'UNIDAD DE NEGOCIO DE ORIGEN',
            'UNIDAD DE NEGOCIO DE DESTINO', 'EMPLEADO', 'TRANSPORTISTA', 'ID INTERNAL',
            'SKU NETSUIT', 'CANTIDAD', 'CENTRO DE COSTO'
        ]
        ws.append(headers)
        for _, row in df.iterrows():
            ws.append([
                row['ID_EXTERNO'], row['FECHA'], row['SUBSIDIARIA'], row['UNIDAD_ORIGEN'],
                row['UNIDAD_DESTINO'], row['EMPLEADO'], row['TRANSPORTISTA'], row['ID_INTERNAL'],
                row['SKU_NETSUIT'], row['CANTIDAD'], row['CENTRO_COSTO']
            ])
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 2, 30)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ------------------------------------------------------------
# PÁGINA DE TRANSFERENCIAS (funcionalidad original)
# ------------------------------------------------------------
def pagina_transferencias():
    st.title("📦 Generador de OT desde movimientos")
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel (con hoja 'Movimientos')", type=["xlsx"])
    if uploaded_file:
        try:
            df_mov = pd.read_excel(uploaded_file, sheet_name="Movimientos")
            st.success("✅ Archivo cargado")
            st.subheader("Vista previa de las primeras filas")
            st.dataframe(df_mov.head(10), width='stretch')

            st.subheader("Tipos de datos de las columnas")
            st.write(df_mov.dtypes)

            debug_container = st.container()
            with debug_container:
                st.subheader("📝 Log de procesamiento")

            if st.button("🚀 Procesar"):
                with st.spinner("Procesando..."):
                    transfers = procesar_movimientos(
                        df_mov, MAPEO_TIENDAS, debug_container)
                    if transfers:
                        excel_data = generar_excel_bytes(transfers)
                        st.success(
                            f"✅ Se generaron {len(transfers)} transferencias")
                        st.download_button(
                            "📥 Descargar Excel", excel_data,
                            f"OT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error(
                            "No se generaron transferencias. Revisa el log de arriba.")
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")
            st.info(
                "Asegúrate de que el archivo tenga una hoja llamada exactamente 'Movimientos'.")

# ------------------------------------------------------------
# PÁGINA DE ÓRDENES DE PEDIDO (NUEVA FUNCIONALIDAD)
# ------------------------------------------------------------
# Mapeo de hojas -> números de tienda que se deben extraer
TIENDAS_POR_HOJA = {
    "GT": [601, 602, 605, 607, 608, 610],
    "SV": [642],
    "HN": [651, 652],
    "CR": [620, 621, 622, 623, 625],
    "PA": [671, 675]
}

# Mapeo para MX (sin PA)
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

@st.cache_data
def cargar_proveedores():
    """Carga el catálogo de proveedores desde data/Proveedores.xlsx."""
    try:
        df = pd.read_excel("data/Proveedores.xlsx", dtype=str)
        # Crear dict nombre -> ID para lookup rápido
        lookup = dict(zip(df["Proveedor"].str.strip(), df["ID_PROVEEDOR"].str.strip()))
        return lookup
    except Exception as e:
        st.error(f"⚠️ No se pudo cargar data/Proveedores.xlsx: {e}")
        return {}

@st.cache_data
def cargar_tiendas():
    """Carga el catálogo de tiendas desde data/Unidad_de_Negocio.xlsx."""
    try:
        df = pd.read_excel("data/Unidad_de_Negocio.xlsx", dtype=str)
        lookup = {}
        for _, row in df.iterrows():
            num = str(row["No. Tienda"]).strip()
            lookup[num] = {
                "unidad_negocio": str(row.get("UNIDAD DE NEGOCIO", "")).strip(),
                "centro_costo":   str(row.get("CENTRO DE COSTO", "")).strip(),
                "subsidiaria":    str(row.get("SUBSIDARIA", "")).strip(),
                "nombre_tienda":  str(row.get("Unidad de Negocio del inventario", "")).strip(),
            }
        return lookup
    except Exception as e:
        st.error(f"⚠️ No se pudo cargar data/Unidad_de_Negocio.xlsx: {e}")
        return {}

def procesar_archivo_pedido_general(uploaded_file, tipo_compra, tiendas_por_hoja, sufijo=""):
    """
    Función genérica para procesar archivo de pedido.
    tiendas_por_hoja: dict con hoja -> lista de tiendas.
    sufijo: se agrega al nombre del archivo CSV (ej: "_MX").
    """
    hoy = datetime.now().strftime("%d/%m/%Y")
    todas_las_filas = []

    lookup_prov   = cargar_proveedores()
    lookup_tienda = cargar_tiendas()

    for hoja, tiendas in tiendas_por_hoja.items():
        try:
            df_raw = pd.read_excel(uploaded_file, sheet_name=hoja,
                                   header=None, dtype=str)
        except Exception as e:
            st.warning(f"No se pudo leer la hoja '{hoja}'. Error: {e}")
            continue

        if len(df_raw) < 14:
            st.warning(f"La hoja '{hoja}' tiene menos de 14 filas, se omite.")
            continue

        header_row = df_raw.iloc[12]
        col_index_map = {}
        for idx, val in enumerate(header_row):
            if pd.isna(val):
                continue
            val_str = str(val).strip()
            if val_str.isdigit():
                col_index_map[int(val_str)] = idx

        tiendas_faltantes = [t for t in tiendas if t not in col_index_map]
        if tiendas_faltantes:
            st.warning(f"En la hoja '{hoja}' faltan las tiendas: {tiendas_faltantes}. Se omitirán.")
            tiendas_a_usar = [t for t in tiendas if t in col_index_map]
        else:
            tiendas_a_usar = tiendas

        for idx_fila in range(13, len(df_raw)):
            fila = df_raw.iloc[idx_fila]
            id_interno = str(fila[1]).strip() if pd.notna(fila[1]) else ""
            if not id_interno or id_interno == "nan":
                continue

            nombre_proveedor = str(fila[0]).strip() if pd.notna(fila[0]) else ""
            sku         = str(fila[2]).strip() if pd.notna(fila[2]) else ""
            descripcion = str(fila[3]).strip() if pd.notna(fila[3]) else ""
            pack        = str(fila[4]).strip() if pd.notna(fila[4]) else ""

            id_proveedor = lookup_prov.get(nombre_proveedor, "#SIN_MATCH")

            for tienda in tiendas_a_usar:
                idx_col = col_index_map[tienda]
                valor_cantidad = str(fila[idx_col]).strip() if pd.notna(fila[idx_col]) else "0"
                try:
                    cantidad_limpia = ''.join(c for c in valor_cantidad if c.isdigit() or c == '.')
                    cantidad = float(cantidad_limpia) if cantidad_limpia else 0.0
                    if cantidad <= 0:
                        continue
                    if cantidad.is_integer():
                        cantidad = int(cantidad)
                except Exception:
                    continue

                tienda_key  = str(tienda)
                datos_tienda = lookup_tienda.get(tienda_key, {})
                unidad_neg  = datos_tienda.get("unidad_negocio", "")
                centro_costo = datos_tienda.get("centro_costo", "")
                subsidiaria  = datos_tienda.get("subsidiaria", "")
                nombre_tienda = datos_tienda.get("nombre_tienda", "")

                # External ID: usa ID de proveedor + tienda + fecha (sin espacios)
                external_id = f"OcBrian{id_proveedor}{tienda}{hoy.replace('/','')}"

                nueva_fila = {
                    "EXTERNAL ID":             external_id,
                    "PROVEEDOR":               id_proveedor,
                    "NOMBRE PROVEDOR":         "",
                    "FECHA":                   hoy,
                    "TIPO DE COMPRA OD":       tipo_compra,
                    "NOTA":                    "",
                    "MONEDA":                  "US Dollar",
                    "UNIDAD DE NEGOCIO":       unidad_neg,
                    "CENTRO DE COSTO":         centro_costo,
                    "SUBSIDIARIA":             subsidiaria,
                    "ARTICULO":                id_interno,
                    "CANTIDAD":                cantidad,
                    "COSTO":                   "",
                    "UNIDAD DE NEGOCIO_2":     unidad_neg,
                    "CENTRO DE COSTO_2":       centro_costo,
                    "validador.tiendanombre":  nombre_tienda,
                    "validador.proveedornombre": nombre_proveedor,
                }
                todas_las_filas.append(nueva_fila)

    if not todas_las_filas:
        st.warning("No se encontraron datos para generar la orden de compra. Verifica el archivo.")
        return pd.DataFrame()

    columnas_orden = [
        "EXTERNAL ID", "PROVEEDOR", "NOMBRE PROVEDOR", "FECHA",
        "TIPO DE COMPRA OD", "NOTA", "MONEDA", "UNIDAD DE NEGOCIO",
        "CENTRO DE COSTO", "SUBSIDIARIA", "ARTICULO", "CANTIDAD",
        "COSTO", "UNIDAD DE NEGOCIO_2", "CENTRO DE COSTO_2",
        "validador.tiendanombre", "validador.proveedornombre",
    ]
    df_final = pd.DataFrame(todas_las_filas)[columnas_orden]
    df_final = df_final.sort_values(by="EXTERNAL ID", ascending=True).reset_index(drop=True)
    return df_final

def pagina_pedidos():
    st.title("📦 Generador de Orden de Compra (estándar)")
    st.markdown("Sube el archivo `Nuevo Análisis V2.xlsx` para generar el archivo de órdenes de compra.")

    separador = st.selectbox(
        "Separador del CSV",
        options=[",", ";"],
        format_func=lambda x: f'Coma  ","' if x == "," else f'Punto y coma  ";"',
        index=0,
        key="sep_pedido"
    )

    tipo_compra = st.selectbox(
        "Tipo de compra",
        options=TIPOS_COMPRA,
        index=0,
        key="tipo_compra_pedido"
    )

    uploaded_file = st.file_uploader(
        "Cargar archivo Excel", type=["xlsx"], key="pedido_uploader")

    if uploaded_file is not None:
        with st.spinner("Procesando archivo..."):
            df_resultado = procesar_archivo_pedido_general(
                uploaded_file, tipo_compra, TIENDAS_POR_HOJA, sufijo=""
            )

        if not df_resultado.empty:
            st.success(f"✅ Procesamiento exitoso. Se generaron {len(df_resultado)} filas para la orden de compra.")
            st.subheader("📊 Vista previa de la orden de compra")
            st.dataframe(df_resultado, use_container_width=True)

            # Generar CSV con el separador elegido
            csv_buffer = io.StringIO()
            encabezados_csv = [
                "EXTERNAL ID", "PROVEEDOR", "NOMBRE PROVEDOR", "FECHA",
                "TIPO DE COMPRA OD", "NOTA", "MONEDA", "UNIDAD DE NEGOCIO",
                "CENTRO DE COSTO", "SUBSIDIARIA", "ARTICULO", "CANTIDAD",
                "COSTO", "UNIDAD DE NEGOCIO", "CENTRO DE COSTO",
                "validador.tiendanombre", "validador.proveedornombre",
            ]
            csv_buffer.write(separador.join(encabezados_csv) + '\n')
            for _, row in df_resultado.iterrows():
                fila_vals = [
                    row["EXTERNAL ID"],   row["PROVEEDOR"],       row["NOMBRE PROVEDOR"],
                    row["FECHA"],         row["TIPO DE COMPRA OD"], row["NOTA"],
                    row["MONEDA"],        row["UNIDAD DE NEGOCIO"], row["CENTRO DE COSTO"],
                    row["SUBSIDIARIA"],   row["ARTICULO"],          row["CANTIDAD"],
                    row["COSTO"],         row["UNIDAD DE NEGOCIO_2"], row["CENTRO DE COSTO_2"],
                    row["validador.tiendanombre"], row["validador.proveedornombre"],
                ]
                if separador == ",":
                    fila_str = separador.join(
                        f'"{v}"' if separador in str(v) else str(v)
                        for v in fila_vals
                    )
                else:
                    fila_str = separador.join(str(v) for v in fila_vals)
                csv_buffer.write(fila_str + '\n')
            csv_data = csv_buffer.getvalue()

            st.download_button(
                label="📥 Descargar CSV",
                data=csv_data,
                file_name=f"orden_compra_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                key="pedido_descarga"
            )
        else:
            st.error("No se pudo generar la orden de compra. Revisa el formato del archivo.")
    else:
        st.info("Por favor, sube un archivo Excel para comenzar.")

def pagina_pedidos_mx():
    st.title("📦 Generador de Orden de Compra (MX)")
    st.markdown("Sube el archivo `Nuevo Análisis V2.xlsx` para generar el archivo de órdenes de compra con hojas MX (GT-MX, SV-MX, HN-MX, CR-MX).")

    separador = st.selectbox(
        "Separador del CSV",
        options=[",", ";"],
        format_func=lambda x: f'Coma  ","' if x == "," else f'Punto y coma  ";"',
        index=0,
        key="sep_pedido_mx"
    )

    tipo_compra = st.selectbox(
        "Tipo de compra",
        options=TIPOS_COMPRA,
        index=0,
        key="tipo_compra_pedido_mx"
    )

    uploaded_file = st.file_uploader(
        "Cargar archivo Excel", type=["xlsx"], key="pedido_uploader_mx")

    if uploaded_file is not None:
        with st.spinner("Procesando archivo..."):
            df_resultado = procesar_archivo_pedido_general(
                uploaded_file, tipo_compra, TIENDAS_POR_HOJA_MX, sufijo="_MX"
            )

        if not df_resultado.empty:
            st.success(f"✅ Procesamiento exitoso. Se generaron {len(df_resultado)} filas para la orden de compra (MX).")
            st.subheader("📊 Vista previa de la orden de compra (MX)")
            st.dataframe(df_resultado, use_container_width=True)

            csv_buffer = io.StringIO()
            encabezados_csv = [
                "EXTERNAL ID", "PROVEEDOR", "NOMBRE PROVEDOR", "FECHA",
                "TIPO DE COMPRA OD", "NOTA", "MONEDA", "UNIDAD DE NEGOCIO",
                "CENTRO DE COSTO", "SUBSIDIARIA", "ARTICULO", "CANTIDAD",
                "COSTO", "UNIDAD DE NEGOCIO", "CENTRO DE COSTO",
                "validador.tiendanombre", "validador.proveedornombre",
            ]
            csv_buffer.write(separador.join(encabezados_csv) + '\n')
            for _, row in df_resultado.iterrows():
                fila_vals = [
                    row["EXTERNAL ID"],   row["PROVEEDOR"],       row["NOMBRE PROVEDOR"],
                    row["FECHA"],         row["TIPO DE COMPRA OD"], row["NOTA"],
                    row["MONEDA"],        row["UNIDAD DE NEGOCIO"], row["CENTRO DE COSTO"],
                    row["SUBSIDIARIA"],   row["ARTICULO"],          row["CANTIDAD"],
                    row["COSTO"],         row["UNIDAD DE NEGOCIO_2"], row["CENTRO DE COSTO_2"],
                    row["validador.tiendanombre"], row["validador.proveedornombre"],
                ]
                if separador == ",":
                    fila_str = separador.join(
                        f'"{v}"' if separador in str(v) else str(v)
                        for v in fila_vals
                    )
                else:
                    fila_str = separador.join(str(v) for v in fila_vals)
                csv_buffer.write(fila_str + '\n')
            csv_data = csv_buffer.getvalue()

            st.download_button(
                label="📥 Descargar CSV (MX)",
                data=csv_data,
                file_name=f"orden_compra_MX_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                key="pedido_descarga_mx"
            )
        else:
            st.error("No se pudo generar la orden de compra (MX). Revisa el formato del archivo.")
    else:
        st.info("Por favor, sube un archivo Excel para comenzar.")

# ------------------------------------------------------------
# INTERFAZ PRINCIPAL CON SELECCIÓN
# ------------------------------------------------------------
st.set_page_config(page_title="Generador de Documentos", layout="wide")

st.sidebar.title("Navegación")
opcion = st.sidebar.radio(
    "¿Qué deseas hacer?",
    ("Formatear Transferencia", "Formatear Orden de Pedido", "Formatear Orden de Pedido (MX)")
)

if opcion == "Formatear Transferencia":
    pagina_transferencias()
elif opcion == "Formatear Orden de Pedido":
    pagina_pedidos()
else:
    pagina_pedidos_mx()
