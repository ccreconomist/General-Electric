import pandas as pd
import warnings
import os
from pathlib import Path

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

ruta_bd_leads = os.path.join(carpeta_destino, "2.BD-Leads.xlsx")
ruta_Leads_MKT = os.path.join(carpeta_destino, "1.Leads_MKT.xlsx")
ruta_CP = os.path.join(carpeta_destino, "3.Clientes Potenciales.xlsx")
ruta_resultado = os.path.join(carpeta_destino, "5.Concentrado Leads.xlsx")

# Crear la carpeta de destino si no existe
Path(carpeta_destino).mkdir(parents=True, exist_ok=True)

# Borrar el archivo '5.Concentrado Leads.xlsx' si ya existe
if os.path.exists(ruta_resultado):
    os.remove(ruta_resultado)

# Leer los archivos Excel
df_bd_leads = pd.read_excel(ruta_bd_leads, engine='openpyxl')
df_Leads_MKT = pd.read_excel(ruta_Leads_MKT, engine='openpyxl')
df_CP = pd.read_excel(ruta_CP, engine='openpyxl')

# Seleccionar las columnas relevantes de cada DataFrame
columnas_bd_leads = ['CORREO', 'NOMBRE', 'CELULAR', 'ESTATUS', 'ESTADO', 'CAMPAÑA', 'CANAL DE PROCEDENCIA', 'FECHA DE CREACIÓn', 'ESTADO', 'ESPECIALIDAD']
columnas_leads_mkt = ['Correo', 'Nombre', 'Celular', 'Canal de Procedencia', 'Fecha-MKT', 'Estado', 'Especialidad']
columnas_CP = ['CORREO', 'NOMBRE', 'CELULAR', 'CANAL DE PROCEDENCIA', 'CREADO', 'ESTADO', 'ASESOR', 'ULTIMO CONTACTO', 'ESTADO', 'ESPECIALIDAD']


# Crear un DataFrame con todos los correos únicos
df_todos_correos = pd.DataFrame({'CORREO': pd.concat([df_bd_leads['CORREO'], df_Leads_MKT['Correo'], df_CP['CORREO']], ignore_index=True).drop_duplicates()})

# Agregar la nueva columna 'ORIGEN' antes de la columna 'CORREO'
df_todos_correos['ORIGEN'] = ''
df_todos_correos.loc[df_todos_correos['CORREO'].isin(df_bd_leads['CORREO']), 'ORIGEN'] += ' BD_LEADS'
df_todos_correos.loc[df_todos_correos['CORREO'].isin(df_Leads_MKT['Correo']), 'ORIGEN'] += ' LEADS_MKT'
df_todos_correos.loc[df_todos_correos['CORREO'].isin(df_CP['CORREO']), 'ORIGEN'] += ' CP'

# Exportar el DataFrame con ORIGEN a un nuevo archivo Excel (hoja "Sheet1")
df_todos_correos.to_excel(ruta_resultado, sheet_name='Sheet1', index=False, engine='openpyxl')

# Fusionar DataFrames utilizando outer join
df_combinado = pd.merge(df_todos_correos, df_bd_leads[columnas_bd_leads], on='CORREO', how='left')
df_combinado = pd.merge(df_combinado, df_Leads_MKT[columnas_leads_mkt], left_on='CORREO', right_on='Correo', how='left')
df_combinado = pd.merge(df_combinado, df_CP[columnas_CP], left_on='CORREO', right_on='CORREO', how='left')

# Verificar si las columnas existen antes de intentar eliminarlas
columnas_a_eliminar = ['Correo']
columnas_existentes = df_combinado.columns.tolist()

for columna in columnas_a_eliminar:
    if columna in columnas_existentes:
        df_combinado.drop(columna, axis=1, inplace=True)

# Guardar el DataFrame final en el archivo de resultado
df_combinado.to_excel(ruta_resultado, sheet_name='Sheet1', index=False, engine='openpyxl')

