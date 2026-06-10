import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font
from datetime import datetime
import os

# Mapa de subsidiaria (número) a texto oficial (según tus datos)
SUBSIDIARIA_TEXTO = {
    15: "0211 OD GUATEMALA Y COMPAÑIA LIMITADA",
    16: "0213 OD EL SALVADOR LTDA, DE C.V.",
    17: "0216 OD HONDURAS S DE R L",
    18: "0214 OD ERIAL BQ S.A",
    19: "0217 OD PANAMA, S.A.",
}

def load_tiendas(file_path):
    """Carga la hoja Tiendas y crea diccionarios: codigo_tienda -> (unidad_negocio, centro_costos, subsidiaria)"""
    df_tiendas = pd.read_excel(file_path, sheet_name="Tiendas")
    # La columna con el código de tienda (ej 620) está en 'UNIDAD DE NEGOCIO'? No, en tu ejemplo: 
    # en la hoja Tiendas, 'UNIDAD DE NEGOCIO' contiene valores como 82,85, etc. y el nombre de tienda tiene código? 
    # En realidad, necesitamos mapear desde el código que aparece en el encabezado de movimientos (ej 620) 
    # hacia los datos de la hoja Tiendas. En tu hoja Tiendas, la columna 'id' parece ser el identificador numérico de la tienda (ej 620? Veo que 620 aparece en la fila 43? Revisando tu tabla Tiendas:
    # 43 | OD | CR | 620 TIENDA ESCAZU | 620 | 229 | 18
    # Allí 'UNIDAD DE NEGOCIO' es 43, 'id' es 620, 'CentroCostos' es 229, 'Subsidiaria' es 18.
    # Por lo tanto, la clave de búsqueda es la columna 'id' (el número de tienda). 
    # El nombre de la tienda en movimientos incluye ese id: "620 ESCAZU". Extraeremos el id numérico.
    # Construimos diccionario: id_tienda (int) -> (unidad_negocio, centro_costos, subsidiaria)
    mapping = {}
    for _, row in df_tiendas.iterrows():
        id_tienda = row['id']   # el número 620, 621, etc.
        unidad_negocio = row['UNIDAD DE NEGOCIO']
        centro_costos = row['CentroCostos']
        subsidiaria = row['Subsidiaria']
        mapping[id_tienda] = (unidad_negocio, centro_costos, subsidiaria)
    return mapping

def parse_movimientos(file_path, tiendas_mapping):
    """Lee hoja Movimientos y genera lista de transferencias (cada una es un dict)"""
    df_mov = pd.read_excel(file_path, sheet_name="Movimientos")
    # Identificar columnas de tiendas: aquellas cuyo nombre contiene un número al inicio (ej "620 ESCAZU")
    tienda_columns = []
    for col in df_mov.columns:
        parts = col.split(' ', 1)
        if parts[0].isdigit():
            tienda_columns.append(col)
    
    # Diccionario para recoger todas las transferencias
    all_transfers = []  # list of dict
    
    # Para cada fila (SKU)
    for idx, row in df_mov.iterrows():
        sku = row['SKU']
        id_interno = row['ID interno']
        # Obtener cantidades por tienda
        origenes = []  # list of (id_tienda, cantidad_abs)
        destinos = []
        for col in tienda_columns:
            cantidad = row[col]
            if cantidad == 0 or pd.isna(cantidad):
                continue
            # Extraer id tienda del nombre de columna
            id_tienda = int(col.split(' ')[0])
            if cantidad < 0:
                origenes.append((id_tienda, -cantidad))  # guardamos positivo
            elif cantidad > 0:
                destinos.append((id_tienda, cantidad))
            else:
                pass
        
        # Validar balance
        total_origen = sum(c for _, c in origenes)
        total_destino = sum(c for _, c in destinos)
        if abs(total_origen - total_destino) > 0.001:
            print(f"Error: SKU {sku} no balanceado (origen {total_origen}, destino {total_destino}) - omitido")
            continue
        
        # Emparejamiento greedy
        # Trabajamos con listas mutables
        orig = [[tid, cant] for tid, cant in origenes]  # lista de [id, cantidad]
        dest = [[tid, cant] for tid, cant in destinos]
        
        while orig and dest:
            # ordenar por cantidad descendente
            orig.sort(key=lambda x: x[1], reverse=True)
            dest.sort(key=lambda x: x[1], reverse=True)
            o_id, o_cant = orig[0]
            d_id, d_cant = dest[0]
            transfer = min(o_cant, d_cant)
            # Obtener datos de la tienda origen (para unidad de negocio, centro costo, subsidiaria)
            if o_id not in tiendas_mapping:
                print(f"Error: id_tienda {o_id} no encontrado en Tiendas (SKU {sku})")
                continue
            if d_id not in tiendas_mapping:
                print(f"Error: id_tienda {d_id} no encontrado en Tiendas (SKU {sku})")
                continue
            unidad_origen, centro_origen, subsid_origen = tiendas_mapping[o_id]
            _, _, subsid_destino = tiendas_mapping[d_id]
            # Validar mismo país
            if subsid_origen != subsid_destino:
                print(f"Error: Transferencia entre países diferentes: tienda {o_id} (subsidiaria {subsid_origen}) -> {d_id} (subsidiaria {subsid_destino}) - SKU {sku} omitido")
                continue
            
            transfer_dict = {
                'ID_EXTERNO': None,  # se calculará después
                'FECHA': datetime.now().date(),  # fecha actual (para luego convertir a número serial)
                'SUBSIDIARIA': SUBSIDIARIA_TEXTO.get(subsid_origen, f"SUBSIDIARIA_{subsid_origen}"),
                'UNIDAD_ORIGEN': unidad_origen,
                'UNIDAD_DESTINO': tiendas_mapping[d_id][0],
                'EMPLEADO': '',
                'TRANSPORTISTA': 'TRANSPORTE PROPIO',
                'ID_INTERNAL': id_interno,
                'SKU_NETSUIT': sku,
                'CANTIDAD': transfer,
                'CENTRO_COSTO': centro_origen,
                'subsidiaria_num': subsid_origen  # para agrupar luego
            }
            all_transfers.append(transfer_dict)
            # Actualizar cantidades
            orig[0][1] -= transfer
            if orig[0][1] == 0:
                orig.pop(0)
            dest[0][1] -= transfer
            if dest[0][1] == 0:
                dest.pop(0)
    
    return all_transfers

def fecha_serial_excel(fecha_date):
    """Convierte datetime.date a número serial de Excel (días desde 1899-12-30)"""
    base = datetime(1899, 12, 30).date()
    delta = fecha_date - base
    return delta.days

def generar_excel_salida(transfers, output_path):
    """Genera archivo Excel con una hoja por subsidiaria, formato OT"""
    # Agrupar por subsidiaria_num
    grupos = {}
    for t in transfers:
        sub_num = t.pop('subsidiaria_num')  # lo usamos solo para agrupar
        grupos.setdefault(sub_num, []).append(t)
    
    # Crear libro
    wb = Workbook()
    # Eliminar hoja por defecto
    wb.remove(wb.active)
    
    # Mapa de subsidiaria_num a nombre de hoja (usando tu nomenclatura)
    hoja_nombres = {
        15: "OT-GT",
        16: "OT-SV",
        17: "OT-HN",
        18: "OT-CR",
        19: "OT-PA",
    }
    
    for sub_num, lista in grupos.items():
        # Crear DataFrame
        df = pd.DataFrame(lista)
        # Reordenar columnas según el formato original
        columnas_orden = ['ID_EXTERNO', 'FECHA', 'SUBSIDIARIA', 'UNIDAD_ORIGEN', 'UNIDAD_DESTINO',
                          'EMPLEADO', 'TRANSPORTISTA', 'ID_INTERNAL', 'SKU_NETSUIT', 'CANTIDAD', 'CENTRO_COSTO']
        df = df[columnas_orden]
        # Calcular ID_EXTERNO y FECHA serial
        # Obtener fecha de hoy (todas las filas tienen misma fecha)
        hoy = datetime.now().date()
        fecha_serial = fecha_serial_excel(hoy)
        df['ID_EXTERNO'] = df.apply(lambda row: f"OT{fecha_serial}{row['UNIDAD_ORIGEN']}{row['UNIDAD_DESTINO']}", axis=1)
        df['FECHA'] = fecha_serial  # número serial de Excel
        # Reemplazar en SUBSIDIARIA el texto correcto (ya lo tenemos de antes, pero aseguramos)
        # Convertir a tipo objeto para evitar problemas con enteros
        df = df.astype({
            'ID_EXTERNO': str,
            'FECHA': int,
            'SUBSIDIARIA': str,
            'UNIDAD_ORIGEN': int,
            'UNIDAD_DESTINO': int,
            'EMPLEADO': str,
            'TRANSPORTISTA': str,
            'ID_INTERNAL': int,
            'SKU_NETSUIT': str,
            'CANTIDAD': int,
            'CENTRO_COSTO': int
        })
        
        # Crear hoja
        sheet_name = hoja_nombres.get(sub_num, f"OT-{sub_num}")
        ws = wb.create_sheet(title=sheet_name)
        # Escribir encabezados (en mayúsculas como en original)
        headers = ['ID EXTERNO', 'FECHA', 'SUBSIDIARIA', 'UNIDAD DE NEGOCIO DE ORIGEN',
                   'UNIDAD DE NEGOCIO DE DESTINO', 'EMPLEADO', 'TRANSPORTISTA', 'ID INTERNAL',
                   'SKU NETSUIT', 'CANTIDAD', 'CENTRO DE COSTO']
        ws.append(headers)
        # Escribir datos
        for _, row in df.iterrows():
            ws.append([
                row['ID_EXTERNO'], row['FECHA'], row['SUBSIDIARIA'], row['UNIDAD_ORIGEN'],
                row['UNIDAD_DESTINO'], row['EMPLEADO'], row['TRANSPORTISTA'], row['ID_INTERNAL'],
                row['SKU_NETSUIT'], row['CANTIDAD'], row['CENTRO_COSTO']
            ])
        # Ajustar ancho de columnas (opcional)
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
    
    wb.save(output_path)
    print(f"Archivo generado: {output_path}")

def main():
    input_file = input("Ingrese la ruta del archivo Excel de entrada (debe contener hojas 'Movimientos' y 'Tiendas'): ").strip()
    if not os.path.exists(input_file):
        print("Archivo no encontrado.")
        return
    output_file = f"OT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    print("Cargando mapeo de tiendas...")
    tiendas_map = load_tiendas(input_file)
    print("Procesando movimientos...")
    transfers = parse_movimientos(input_file, tiendas_map)
    if not transfers:
        print("No se generaron transferencias. Verifique los datos.")
        return
    print(f"Se generaron {len(transfers)} transferencias.")
    generar_excel_salida(transfers, output_file)

if __name__ == "__main__":
    main()