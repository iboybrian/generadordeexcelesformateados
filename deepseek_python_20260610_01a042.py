import streamlit as st
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import io

# ------------------------------------------------------------
# MAPEO DE TIENDAS (código del encabezado → (unidad_negocio, centro_costo, subsidiaria_num))
# Agrega aquí todas las tiendas que puedan aparecer en tus movimientos.
# ------------------------------------------------------------
MAPEO_TIENDAS = {
    601: (82, 241, 15),   # Guatemala
    602: (85, 242, 15),
    605: (91, 244, 15),
    607: (97, 246, 15),
    608: (100, 247, 15),
    610: (106, 249, 15),
    # Puedes añadir más, por ejemplo:
    # 620: (43, 229, 18),   # Costa Rica
    # 621: (46, 233, 18),
}

# Texto oficial de cada subsidiaria según el número
SUBSIDIARIA_TEXTO = {
    15: "0211 OD GUATEMALA Y COMPAÑIA LIMITADA",
    16: "0213 OD EL SALVADOR LTDA, DE C.V.",
    17: "0216 OD HONDURAS S DE R L",
    18: "0214 OD ERIAL BQ S.A",
    19: "0217 OD PANAMA, S.A.",
}

# Nombres de las hojas de salida por subsidiaria
NOMBRE_HOJA = {15: "OT-GT", 16: "OT-SV", 17: "OT-HN", 18: "OT-CR", 19: "OT-PA"}

# ------------------------------------------------------------
# FUNCIONES AUXILIARES
# ------------------------------------------------------------
def fecha_serial_excel(fecha_date):
    """Convierte datetime.date a número serial de Excel (días desde 1899-12-30)"""
    base = datetime(1899, 12, 30).date()
    return (fecha_date - base).days

def procesar_movimientos(df_mov, mapeo):
    """
    Procesa el DataFrame de movimientos y devuelve una lista de transferencias.
    Cada transferencia es un diccionario con los campos necesarios.
    """
    # Identificar columnas de tienda: aquellas cuyo nombre comienza con un número
    tienda_cols = []
    for col in df_mov.columns:
        try:
            int(col.split()[0])
            tienda_cols.append(col)
        except:
            pass

    transfers = []
    for idx, row in df_mov.iterrows():
        sku = row['SKU']
        id_interno = row['ID interno']
        origenes = []   # lista de [id_tienda, cantidad_abs]
        destinos = []

        for col in tienda_cols:
            cantidad = row[col]
            if pd.isna(cantidad) or cantidad == 0:
                continue
            id_tienda = int(col.split()[0])
            if cantidad < 0:
                origenes.append([id_tienda, -cantidad])
            else:
                destinos.append([id_tienda, cantidad])

        # Validar balance
        total_origen = sum(c for _, c in origenes)
        total_destino = sum(c for _, c in destinos)
        if abs(total_origen - total_destino) > 0.001:
            st.warning(f"SKU {sku} no balanceado (origen {total_origen}, destino {total_destino}) - omitido")
            continue

        # Emparejamiento greedy
        orig = origenes.copy()
        dest = destinos.copy()
        while orig and dest:
            orig.sort(key=lambda x: x[1], reverse=True)
            dest.sort(key=lambda x: x[1], reverse=True)
            o_id, o_cant = orig[0]
            d_id, d_cant = dest[0]
            transferir = min(o_cant, d_cant)

            if o_id not in mapeo:
                st.error(f"Tienda origen {o_id} no está en el mapeo (SKU {sku})")
                break
            if d_id not in mapeo:
                st.error(f"Tienda destino {d_id} no está en el mapeo (SKU {sku})")
                break

            unidad_origen, centro_origen, sub_origen = mapeo[o_id]
            unidad_destino, _, sub_destino = mapeo[d_id]

            # Validar mismo país
            if sub_origen != sub_destino:
                st.error(f"Transferencia entre países: {o_id} → {d_id} (SKU {sku}) - omitida")
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

            # Actualizar cantidades restantes
            orig[0][1] -= transferir
            if orig[0][1] == 0:
                orig.pop(0)
            dest[0][1] -= transferir
            if dest[0][1] == 0:
                dest.pop(0)

    return transfers

def generar_excel_bytes(transfers):
    """Recibe la lista de transferencias y devuelve los bytes del archivo Excel."""
    if not transfers:
        return None

    # Agrupar por subsidiaria
    grupos = {}
    for t in transfers:
        sub = t['subsidiaria_num']
        grupos.setdefault(sub, []).append(t)

    wb = Workbook()
    wb.remove(wb.active)  # eliminar hoja por defecto

    hoy = datetime.now().date()
    fecha_serial = fecha_serial_excel(hoy)

    for sub_num, lista in grupos.items():
        df = pd.DataFrame(lista)
        # Calcular columnas adicionales
        df['ID_EXTERNO'] = df.apply(lambda row: f"OT{fecha_serial}{row['UNIDAD_ORIGEN']}{row['UNIDAD_DESTINO']}", axis=1)
        df['FECHA'] = fecha_serial
        df['SUBSIDIARIA'] = SUBSIDIARIA_TEXTO.get(sub_num, f"SUBSIDIARIA_{sub_num}")
        df['EMPLEADO'] = ''
        df['TRANSPORTISTA'] = 'TRANSPORTE PROPIO'

        # Reordenar columnas según el formato deseado
        columnas_orden = [
            'ID_EXTERNO', 'FECHA', 'SUBSIDIARIA', 'UNIDAD_ORIGEN', 'UNIDAD_DESTINO',
            'EMPLEADO', 'TRANSPORTISTA', 'ID_INTERNAL', 'SKU_NETSUIT', 'CANTIDAD', 'CENTRO_COSTO'
        ]
        df = df[columnas_orden]

        # Crear hoja
        sheet_name = NOMBRE_HOJA.get(sub_num, f"OT-{sub_num}")
        ws = wb.create_sheet(title=sheet_name)

        # Encabezados (con espacios, tal como se espera)
        headers = [
            'ID EXTERNO', 'FECHA', 'SUBSIDIARIA', 'UNIDAD DE NEGOCIO DE ORIGEN',
            'UNIDAD DE NEGOCIO DE DESTINO', 'EMPLEADO', 'TRANSPORTISTA', 'ID INTERNAL',
            'SKU NETSUIT', 'CANTIDAD', 'CENTRO DE COSTO'
        ]
        ws.append(headers)

        # Escribir filas
        for _, row in df.iterrows():
            ws.append([
                row['ID_EXTERNO'], row['FECHA'], row['SUBSIDIARIA'], row['UNIDAD_ORIGEN'],
                row['UNIDAD_DESTINO'], row['EMPLEADO'], row['TRANSPORTISTA'], row['ID_INTERNAL'],
                row['SKU_NETSUIT'], row['CANTIDAD'], row['CENTRO_COSTO']
            ])

        # Ajustar ancho de columnas
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

    # Guardar en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ------------------------------------------------------------
# INTERFAZ DE STREAMLIT
# ------------------------------------------------------------
st.set_page_config(page_title="Generador de Órdenes de Transferencia", layout="wide")
st.title("📦 Generador de OT desde movimientos entre tiendas")

st.markdown("""
### Instrucciones
1. Sube un archivo **Excel** que contenga una hoja llamada **`Movimientos`**.
2. La hoja debe tener las columnas: `ID interno`, `SKU`, `Descripcion` (opcional), `UR` (opcional) y luego columnas con nombres como `601 TIENDA X`, `602 TIENDA Y`, etc.
3. Los valores **negativos** = origen, **positivos** = destino.
4. El archivo de salida será un Excel con una hoja por cada subsidiaria (país) involucrada, en el formato de OT que esperas.
""")

uploaded_file = st.file_uploader("Selecciona tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_mov = pd.read_excel(uploaded_file, sheet_name="Movimientos")
        st.success("✅ Archivo cargado correctamente")
        st.subheader("Vista previa de los movimientos (primeras 10 filas)")
        st.dataframe(df_mov.head(10))

        if st.button("🚀 Procesar y generar OT"):
            with st.spinner("Procesando movimientos..."):
                transfers = procesar_movimientos(df_mov, MAPEO_TIENDAS)
                if not transfers:
                    st.error("No se generaron transferencias. Revisa los mensajes de advertencia/error arriba.")
                else:
                    excel_data = generar_excel_bytes(transfers)
                    st.success(f"✅ Se generaron {len(transfers)} transferencias.")
                    st.download_button(
                        label="📥 Descargar archivo Excel",
                        data=excel_data,
                        file_name=f"OT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.info("Asegúrate de que el archivo tenga una hoja llamada exactamente 'Movimientos'.")
