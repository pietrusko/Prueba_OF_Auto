import pandas as pd
from datetime import datetime
import os

# Configurar las rutas de los archivos y el nombre de las hojas
ordenes_file = 'H:/.shortcut-targets-by-id/0B1uIg31RZitwRG4yYWNHYTR6eFE/Control Fabrica/2024/Control_Produccion_2024.xlsm'  # Ruta del archivo de órdenes de fabricación
recetas_file = 'H:/Mi unidad/Producción/2024/Recetas productos 1000LT.xlsx'  # Ruta del archivo de recetas
output_file = 'H:/Mi unidad/Producción/2024/Consumos diarios.xlsx'  # Ruta del archivo de salida especificado
sheet_name_ordenes = 'INF_OrdenFAB'  # Nombre de la hoja específica en órdenes de fabricación
sheet_name_recetas = 'recetas 1000L'  # Nombre de la hoja en el archivo de recetas

# Crear el directorio de salida si no existe
output_dir = os.path.dirname(output_file)
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Obtener la fecha actual y el mes actual
hoy = datetime.now().strftime('%Y-%m-%d')
mes_actual = datetime.now().strftime('%Y-%m')

# Leer el archivo de órdenes de fabricación desde la hoja específica, saltando filas de encabezado
df_ordenes = pd.read_excel(ordenes_file, sheet_name=sheet_name_ordenes, skiprows=8)

# Asignar nombres de columnas relevantes si no están automáticamente configurados
df_ordenes.columns = ['OF', 'Producto', 'Mes Fab', 'Año Fab', 'N° OF', 'Producto2', 'Litros', 'Otro1', 'Otro2', 'Otro3', 'Otro4', 'Otro5', 'Otro6', 'Otro7', 'Otro8', 'Otro9', 'Otro10', 'Otro11', 'Otro12']

# Filtrar filas con valores no válidos en 'Año Fab' y 'Mes Fab'
df_ordenes = df_ordenes[pd.to_numeric(df_ordenes['Año Fab'], errors='coerce').notnull()]
df_ordenes = df_ordenes[pd.to_numeric(df_ordenes['Mes Fab'], errors='coerce').notnull()]

# Convertir la columna de fecha al formato adecuado (suponiendo 'Año Fab' y 'Mes Fab' se combinan para formar una fecha)
df_ordenes['fecha'] = pd.to_datetime(df_ordenes['Año Fab'].astype(int).astype(str) + '-' + df_ordenes['Mes Fab'].astype(int).astype(str).str.zfill(2) + '-01')

# Filtrar las órdenes del día actual
ordenes_hoy = df_ordenes[df_ordenes['fecha'] == hoy]

# Leer el archivo de recetas desde la hoja específica
df_recetas = pd.read_excel(recetas_file, sheet_name=sheet_name_recetas)

# Procesar cada producto por separado
productos = ordenes_hoy['Producto'].unique()

# Leer la hoja 'Hoja1' del archivo de consumos diarios para obtener la estructura
df_estructura = pd.read_excel(output_file, sheet_name='Hoja1')

# Diccionario para acumular el consumo total diario de todos los productos
consumo_total_diario = {}

# Crear una estructura similar a la de 'Hoja1' para el output diario
df_output = df_estructura.copy()
df_output[hoy] = 0  # Añadir la columna para el día actual

with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    for producto in productos:
        # Filtrar las órdenes del producto actual
        ordenes_producto = ordenes_hoy[ordenes_hoy['Producto'] == producto]

        # Calcular el total de litros fabricados del producto actual hoy
        total_litros_producto = ordenes_producto['Litros'].sum()

        # Filtrar la receta del producto actual
        receta_producto = df_recetas[df_recetas['Producto'] == producto]

        if receta_producto.empty:
            print(f"No hay receta disponible para el producto {producto}")
            continue

        # Escalar la receta según los litros fabricados
        consumo_diario_producto = receta_producto.copy()
        for columna in receta_producto.columns[1:]:
            consumo_diario_producto[columna] = receta_producto[columna] * total_litros_producto / 1000

        # Sumar las cantidades de cada columna
        sumas_diarias_producto = consumo_diario_producto.sum(numeric_only=True)

        # Agregar el consumo del producto al consumo total diario
        for key, value in sumas_diarias_producto.items():
            if key in df_output['DIA'].values:
                df_output.loc[df_output['DIA'] == key, hoy] += value

    # Nombre de la hoja para el consumo total diario
    sheet_name_total = f'Total_{mes_actual}'

    # Si el archivo de destino no existe, crear uno nuevo con el consumo total diario
    try:
        df_destino_total = pd.read_excel(output_file, sheet_name=sheet_name_total, index_col=0)
        # Agregar el consumo total diario al archivo existente
        df_actualizado_total = pd.concat([df_destino_total, df_output.set_index('DIA')[hoy]], axis=1)
        df_actualizado_total.to_excel(writer, sheet_name=sheet_name_total)
    except ValueError:
        # Si la hoja no existe, crear una nueva
        df_output.set_index('DIA')[hoy].to_excel(writer, sheet_name=sheet_name_total)

    # Guardar el df_output en una hoja para el día actual
    sheet_name_diario = f'Diario_{hoy}'
    df_output.to_excel(writer, sheet_name=sheet_name_diario, index=False)

print(f'Consumo diario de ingredientes para el {hoy} agregado al archivo {output_file}')

