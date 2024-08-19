import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\PS"

ruta_bd_leads = os.path.join(carpeta_destino, "2.BD-Leads.xlsx")
ruta_Leads_MKT = os.path.join(carpeta_destino, "1.Leads_MKT.xlsx")
ruta_CP = os.path.join(carpeta_destino, "3.Clientes Potenciales.xlsx")
ruta_resultado_cp = os.path.join(carpeta_destino, "5.Concentrado Leads.xlsx")

# Definir las columnas a usar
columnas_bd_leads = ['CORREO',	'FOLIO',	'ESTATUS', 'NOMBRE',	'FECHA DE CONTACTO',	'COICIDENCIA CLIENTE POTENCIAL',	'LEAD ASOCIADO',	'ZONA',	'CELULAR',	'CANAL DE PROCEDENCIA',	'TIPO DE CAMPAÑA',	'ORIGEN PROSPECTO',	'ASESOR',	'CLIENTE POTENCIAL',	'FECHA DE ASIGNACIÓN',	'OPORTUNIDADES',	'ESPECIALIDAD',	'CIUDAD',	'ESTADO',	'CAMPAÑA',	'COMENTARIOS',	'FECHA DE CREACIÓn',	'ULTIMO CONTACTO','FRECUENCIA']
columnas_leads_mkt = ['Correo',	'Procedencia',	'Fecha-MKT','Nombre',	'Especialidad',	'Celular',	'Ciudad',	'Estado','Campaña ',	'NOTAS',	'FRECUENCIA_MKT']
columnas_CP = ['CORREO','FOLIO','Cliente potencial','ULTIMO CONTACTO',	'CELULAR',	'ESPECIALIDAD',	'CANAL DE PROCEDENCIA',	'EVENTO/CAMPAÑA',	'ZONA',	'ESTADO',	'CIUDAD',	'ASESOR',	'CREADO-CP', 'FRECUENCIA_CP']

# Leer los archivos Excel con un manejo de errores para columnas ausentes
def leer_excel(ruta, columnas):
    try:
        return pd.read_excel(ruta, engine='openpyxl', usecols=columnas)
    except ValueError as e:
        # Identificar las columnas que no están en el archivo y quitarlas de la lista
        cols_not_found = [col for col in columnas if col not in pd.read_excel(ruta, engine='openpyxl').columns]
        print(f"Columnas no encontradas en {ruta}: {cols_not_found}")
        columnas_actualizadas = [col for col in columnas if col in pd.read_excel(ruta, engine='openpyxl').columns]
        return pd.read_excel(ruta, engine='openpyxl', usecols=columnas_actualizadas)

df_bd_leads = leer_excel(ruta_bd_leads, columnas_bd_leads)
df_Leads_MKT = leer_excel(ruta_Leads_MKT, columnas_leads_mkt)
df_CP = leer_excel(ruta_CP, columnas_CP)

# Renombrar columna 'Correo' de Leads_MKT para unificar
df_Leads_MKT.rename(columns={'Correo': 'CORREO'}, inplace=True)

# Concatenar todos los DataFrames
df_todos = pd.concat([df_bd_leads, df_Leads_MKT, df_CP], ignore_index=True)

# Eliminar duplicados basados en la columna 'CORREO'
df_todos = df_todos.drop_duplicates(subset=['CORREO'], keep='first')

# Inicializar columna 'ORIGEN'
df_todos['ORIGEN'] = ''

# Marcar el origen de los correos en la columna 'ORIGEN'
df_todos.loc[df_todos['CORREO'].isin(df_bd_leads['CORREO']), 'ORIGEN'] += ' BD_LEADS'
df_todos.loc[df_todos['CORREO'].isin(df_Leads_MKT['CORREO']), 'ORIGEN'] += ' LEADS_MKT'
df_todos.loc[df_todos['CORREO'].isin(df_CP['CORREO']), 'ORIGEN'] += ' CP'

# Unir los DataFrames en base a la columna 'CORREO'
df_merged = df_bd_leads.merge(df_Leads_MKT, on='CORREO', how='outer', suffixes=('_BD', '_MKT')).merge(df_CP, on='CORREO', how='outer', suffixes=('', '_CP'))

# Sumar las columnas de 'FRECUENCIA'
df_merged['FRECUENCIA_TOTAL'] = df_merged[['FRECUENCIA', 'FRECUENCIA_MKT', 'FRECUENCIA_CP']].sum(axis=1, skipna=True)

# Agregar la columna 'ORIGEN' al DataFrame final
df_merged = df_merged.merge(df_todos[['CORREO', 'ORIGEN']], on='CORREO', how='left')

# Consolidar filas duplicadas
df_final = df_merged.groupby('CORREO').agg(lambda x: x.dropna().iloc[0] if x.dropna().any() else '').reset_index()

# Guardar el DataFrame resultante en un nuevo archivo Excel
df_final.to_excel(ruta_resultado_cp, sheet_name='Sheet1', index=False, engine='openpyxl')
