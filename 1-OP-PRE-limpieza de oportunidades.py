import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

# Rutas del archivo de oportunidad
ruta_oportunidad = os.path.join(carpeta_destino, "8.Seguimiento OP.xlsx")

# Leer la hoja de oportunidad
df_oportunidades = pd.read_excel(ruta_oportunidad, engine='openpyxl')

# Función para pivotar los DataFrames según las fechas
def pivot_data(df, date_col, cols_to_keep):
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.sort_values(by=date_col)
    df['sequence'] = df.groupby('Cliente potencial').cumcount() + 1
    df_pivoted = df.pivot(index='Cliente potencial', columns='sequence', values=cols_to_keep)
    df_pivoted.columns = [f"{col}_{num}" for col, num in df_pivoted.columns]
    return df_pivoted.reset_index()

# Procesar y pivotar df_oportunidades
cols_oportunidades = ['Actualizado1', 'Oportunidad', 'Tipo de seguimiento', 'Fecha', 'Estatus']
pivoted_oportunidades = pivot_data(df_oportunidades, 'Actualizado1', cols_oportunidades)

# Definir la ruta de destino para el archivo Excel
ruta_destino_excel = os.path.join(carpeta_destino, "resultado_seguimientos_op.xlsx")

# Exportar el DataFrame 'pivoted_oportunidades' a un archivo Excel
pivoted_oportunidades.to_excel(ruta_destino_excel, index=False)

# Confirmación de la exportación
print(f"El archivo Excel se ha exportado exitosamente a: {ruta_destino_excel}")
