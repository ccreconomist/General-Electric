import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\PS"

# Rutas del archivo de oportunidad
ruta_oportunidad = os.path.join(carpeta_destino, "7.Seguimientos.xlsx")

# Leer la hoja de oportunidad
df_oportunidades = pd.read_excel(ruta_oportunidad, engine='openpyxl')

# Comprobar si las columnas necesarias están en el DataFrame
required_columns = ['Folio',	'Cliente potencial',	'Asesor',	'Tipo',	'fecha de contacto-Agenda',	'Estatus',	'Resultado de actividad',	'Fecha Real']
missing_columns = [col for col in required_columns if col not in df_oportunidades.columns]

if missing_columns:
    raise KeyError(f"Las siguientes columnas faltan en el DataFrame: {missing_columns}")

# Función para pivotar los DataFrames según las fechas
def pivot_data(df, date_col, cols_to_keep):
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.sort_values(by=date_col)
    df['sequence'] = df.groupby('Cliente potencial').cumcount() + 1
    df_pivoted = df.pivot(index='Cliente potencial', columns='sequence', values=cols_to_keep)
    df_pivoted.columns = [f"{col}_{num}" for col, num in df_pivoted.columns]
    return df_pivoted.reset_index()

# Procesar y pivotar df_oportunidades
cols_oportunidades = ['Fecha Real', 'Resultado de actividad', 'Estatus', 'fecha de contacto-Agenda', 'Tipo', 'Asesor']
pivoted_oportunidades = pivot_data(df_oportunidades, 'Fecha Real', cols_oportunidades)

# Definir la ruta de destino para el archivo Excel
ruta_destino_excel = os.path.join(carpeta_destino, "resultado_seguimientos_comercial.xlsx")

# Exportar el DataFrame 'pivoted_oportunidades' a un archivo Excel
pivoted_oportunidades.to_excel(ruta_destino_excel, index=False)

# Confirmación de la exportación
print(f"El archivo Excel se ha exportado exitosamente a: {ruta_destino_excel}")


#una vez que se haga lo anterior, con ese mismo archivo de excel resultado_seguimientos_comercial hay que hacer un resumen en otra hoja de excel donde por cliente potencial unico,
#en una fila llamada Primera tomando la fecha mas antigua del mismo cliente potencial de la Fecha Real, y otra columna donde tome la fecha mas reciente de ese cliente potencial,
#en dado caso que sea la misma fecha o que solo haya una fecha, se pone la misma fecha en ambas columnas, luego, otra columna donde seal primer status de acuerdo a las fecha real del primero estatus y el ultimo estatus, y si solo hay uno dejar ese entonces es columna 1 y columna 2
#