import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

# Rutas de los archivos
ruta_OP = os.path.join(carpeta_destino, "6.Oportunidades.xlsx")
ruta_SG = os.path.join(carpeta_destino, "resultado_seguimientos_op.xlsx")

# Leer los archivos de oportunidades y seguimientos
df_op = pd.read_excel(ruta_OP, engine='openpyxl')
df_sg = pd.read_excel(ruta_SG, engine='openpyxl')

# Función para pivotar los DataFrames según las fechas
def pivot_data(df, date_col, cols_to_keep):
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.sort_values(by=date_col)
    df['sequence'] = df.groupby('Cliente potencial').cumcount() + 1
    df_pivoted = df.pivot(index='Cliente potencial', columns='sequence', values=cols_to_keep)
    df_pivoted.columns = [f"{col}_{num}" for col, num in df_pivoted.columns]
    return df_pivoted.reset_index()

# Procesar y pivotar df_op
cols_op = ['Actualizado', 'Concepto Venta', 'Estatus de oportunidad', 'Asesor', 'Origen de oportunidad']
pivoted_op = pivot_data(df_op, 'Actualizado', cols_op)

# Unir las tablas pivotadas por 'Cliente potencial'
resultado_final = pd.merge(df_sg, pivoted_op, on='Cliente potencial', how='outer')

# Definir la ruta de destino para el archivo Excel
ruta_destino_excel = os.path.join(carpeta_destino, "resultado_final.xlsx")

# Exportar el DataFrame 'resultado_final' a un archivo Excel
resultado_final.to_excel(ruta_destino_excel, index=False)

# Confirmación de la exportación
print(f"El archivo Excel se ha exportado exitosamente a: {ruta_destino_excel}")
