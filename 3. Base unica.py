import pandas as pd
import numpy as np
import warnings
import os
from pathlib import Path
from datetime import datetime

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\PS"

ruta_bd_leads = os.path.join(carpeta_destino, "2.BD-Leads.xlsx")
ruta_Leads_MKT = os.path.join(carpeta_destino, "1.Leads_MKT.xlsx")
ruta_CP = os.path.join(carpeta_destino, "3.Clientes Potenciales.xlsx")
ruta_resultado = os.path.join(carpeta_destino, "5.Concentrado Leads.xlsx")

# Leer los archivos Excel
df_bd_leads = pd.read_excel(ruta_bd_leads, engine='openpyxl', usecols=['NOMBRE',	'CORREO',	'CELULAR',	'ESTATUS',	'ESTADO',	'CAMPAÑA',	'CANAL DE PROCEDENCIA',	'FECHA DE CREACIÓn',	'ESPECIALIDAD',	'FOLIO',	'TIPO DE LEAD',	'FECHA DE CONTACTO',	'COICIDENCIA CLIENTE POTENCIAL',	'LEAD ASOCIADO',	'ZONA',	'TIPO DE CAMPAÑA',	'ORIGEN PROSPECTO',	'ASESOR',	'CLIENTE POTENCIAL',	'FECHA DE ASIGNACIÓN',	'OPORTUNIDADES',	'CIUDAD',	'COMENTARIOS',	'ULTIMO CONTACTO'])
df_Leads_MKT = pd.read_excel(ruta_Leads_MKT, engine='openpyxl', usecols=['Correo', 'Nombre', 'Fecha-MKT', 'Campaña', 'Canal de Procedencia', 'Estado', 'Celular', 'Especialidad'])
df_CP = pd.read_excel(ruta_CP, engine='openpyxl', usecols=['FOLIO', 'Cliente potencial', 'ULTIMO CONTACTO', 'CORREO', 'CELULAR', 'ESPECIALIDAD', 'CANAL DE PROCEDENCIA', 'EVENTO/CAMPAÑA', 'ZONA', 'ESTADO', 'CIUDAD', 'ASESOR', 'CREADO-CP', 'Fecha Real', 'Creado-OP', 'CREADO.CP', 'Encontrado_seguimientos', 'Asesor', 'Tipo', 'fecha de contacto-Agenda', 'Estatus', 'Resultado de actividad', 'Oportunidad', 'Tipo de seguimiento', 'Fecha', 'Estatus_oportunidad', 'Resultado de actividad_oportunidad', 'Creado por', 'Actualizado por', 'Actualizado1', 'Año', 'Mes', 'Asesor_oportunidad', 'Estatus de oportunidad', 'Concepto Venta', 'Transductores y/o SW', 'Celular', 'Teléfono', 'Estado', 'Origen de oportunidad', 'Congreso y evento', 'Campaña de Facebook', 'Titulo', 'Monto', '% Conversión', 'Fecha Prob.Cierre', 'Tipo de campaña FB', 'Campaña Comercial', 'Tipo de campaña Google', 'Motivos de rechazo', 'Campaña de Google', 'Creado por_oportunidad', 'Actualizado por_oportunidad', 'Actualizado'])

# Limpiar columnas de texto
df_bd_leads['NOMBRE'] = df_bd_leads['NOMBRE'].str.strip().replace('', np.nan)
df_Leads_MKT['Nombre'] = df_Leads_MKT['Nombre'].str.strip().replace('', np.nan)
df_CP['Cliente potencial'] = df_CP['Cliente potencial'].str.strip().replace('', np.nan)

# Localizar las columnas de fechas y calcular los días transcurridos
hoy = datetime.now()

df_bd_leads['FECHA DE CREACIÓn'] = pd.to_datetime(df_bd_leads['FECHA DE CREACIÓn'], errors='coerce')
df_Leads_MKT['Fecha-MKT'] = pd.to_datetime(df_Leads_MKT['Fecha-MKT'], errors='coerce')
df_CP['CREADO-CP'] = pd.to_datetime(df_CP['CREADO-CP'], errors='coerce')

# Remove duplicates based on the specified columns
df_bd_leads = df_bd_leads.drop_duplicates(['CORREO', 'FECHA DE CREACIÓn'])
df_Leads_MKT = df_Leads_MKT.drop_duplicates(['Correo', 'Fecha-MKT'])
df_CP = df_CP.drop_duplicates(['CORREO', 'CREADO-CP'])

# Calculate the number of days since creation
df_bd_leads['DÍAS DESDE CREACIÓN'] = (hoy - df_bd_leads['FECHA DE CREACIÓn']).dt.days
df_Leads_MKT['DÍAS DESDE CREACIÓN'] = (hoy - df_Leads_MKT['Fecha-MKT']).dt.days
df_CP['DÍAS DESDE CREACIÓN'] = (hoy - df_CP['CREADO-CP']).dt.days

# Definir las categorías de días transcurridos
categorias_dias = {
    (31, 60): '31 a 60 Dias',
    (181, np.inf): 'mayor a 181 Dias',
    (91, 120): '91 a 120 Dias',
    (0, 30): '0 a 30 Dias',
    (121, 150): '121 a 150 Dias',
    (61, 90): '61 a 90 Dias',
    (151, 180): '151 a 180 Dias'
}

# Función para aplicar las categorías
def categorizar_dias(dias):
    for rango, categoria in categorias_dias.items():
        if rango[0] <= dias <= rango[1]:
            return categoria
    return 'Sin categoría'

# Aplicar las categorías a las columnas de días transcurridos
df_bd_leads['CATEGORÍA DÍAS'] = df_bd_leads['DÍAS DESDE CREACIÓN'].apply(categorizar_dias)
df_Leads_MKT['CATEGORÍA DÍAS'] = df_Leads_MKT['DÍAS DESDE CREACIÓN'].apply(categorizar_dias)
df_CP['CATEGORÍA DÍAS'] = df_CP['DÍAS DESDE CREACIÓN'].apply(categorizar_dias)

# Calcular días desde creación para 'Fecha Real' en df_CP y aplicar categorías
df_CP['Fecha Real'] = pd.to_datetime(df_CP['Fecha Real'], errors='coerce')
df_CP['DÍAS DESDE CREACIÓN FECHA REAL'] = (hoy - df_CP['Fecha Real']).dt.days
df_CP['CATEGORÍA DÍAS FECHA REAL'] = df_CP['DÍAS DESDE CREACIÓN FECHA REAL'].apply(categorizar_dias)

# Seleccionar las columnas relevantes de cada DataFrame
columnas_bd_leads = ['NOMBRE',	'CORREO',	'CELULAR',	'ESTATUS',	'ESTADO',	'CAMPAÑA',	'CANAL DE PROCEDENCIA',	'FECHA DE CREACIÓn',	'ESPECIALIDAD',	'FOLIO',	'TIPO DE LEAD',	'FECHA DE CONTACTO',	'COICIDENCIA CLIENTE POTENCIAL',	'LEAD ASOCIADO',	'ZONA',	'TIPO DE CAMPAÑA',	'ORIGEN PROSPECTO',	'ASESOR',	'CLIENTE POTENCIAL',	'FECHA DE ASIGNACIÓN',	'OPORTUNIDADES',	'CIUDAD',	'COMENTARIOS',	'ULTIMO CONTACTO', 'DÍAS DESDE CREACIÓN']
columnas_leads_mkt = ['Correo', 'Nombre', 'Fecha-MKT', 'Campaña', 'Canal de Procedencia', 'Estado', 'Celular', 'Especialidad', 'DÍAS DESDE CREACIÓN']
columnas_CP = ['FOLIO', 'Cliente potencial', 'ULTIMO CONTACTO', 'CORREO', 'CELULAR', 'ESPECIALIDAD', 'CANAL DE PROCEDENCIA', 'EVENTO/CAMPAÑA', 'ZONA', 'ESTADO', 'CIUDAD', 'ASESOR', 'CREADO-CP', 'Fecha Real', 'Creado-OP', 'CREADO.CP', 'Encontrado_seguimientos', 'Asesor', 'Tipo', 'fecha de contacto-Agenda', 'Estatus', 'Resultado de actividad', 'Oportunidad', 'Tipo de seguimiento', 'Fecha', 'Estatus_oportunidad', 'Resultado de actividad_oportunidad', 'Creado por', 'Actualizado por', 'Actualizado1', 'Año', 'Mes', 'Asesor_oportunidad', 'Estatus de oportunidad', 'Concepto Venta', 'Transductores y/o SW', 'Celular', 'Teléfono', 'Estado', 'Origen de oportunidad', 'Congreso y evento', 'Campaña de Facebook', 'Titulo', 'Monto', '% Conversión', 'Fecha Prob.Cierre', 'Tipo de campaña FB', 'Campaña Comercial', 'Tipo de campaña Google', 'Motivos de rechazo', 'Campaña de Google', 'Creado por_oportunidad', 'Actualizado por_oportunidad', 'Actualizado']

# Crear un DataFrame con todos los correos únicos
df_todos_correos = pd.DataFrame({'CORREO': pd.concat([df_bd_leads['CORREO'], df_Leads_MKT['Correo'], df_CP['CORREO']], ignore_index=True).drop_duplicates()})

# Agregar la nueva columna 'ORIGEN' antes de la columna 'CORREO'
df_todos_correos.insert(0, 'ORIGEN', '')

# Marcar el origen de los correos en la columna 'ORIGEN'
df_todos_correos.loc[df_todos_correos['CORREO'].isin(df_bd_leads['CORREO']), 'ORIGEN'] += ' BD_LEADS'
df_todos_correos.loc[df_todos_correos['CORREO'].isin(df_Leads_MKT['Correo']), 'ORIGEN'] += ' LEADS_MKT'
df_todos_correos.loc[df_todos_correos['CORREO'].isin(df_CP['CORREO']), 'ORIGEN'] += ' CP'

# Exportar el DataFrame con ORIGEN a un nuevo archivo Excel (hoja "Sheet1")
df_todos_correos.to_excel(ruta_resultado, sheet_name='Sheet1', index=False, engine='openpyxl')

# Fusionar DataFrames utilizando outer join
df_combinado = pd.merge(df_todos_correos, df_bd_leads[columnas_bd_leads], on='CORREO', how='left', indicator='_merge_BD')
df_combinado = pd.merge(df_combinado, df_Leads_MKT[columnas_leads_mkt], left_on='CORREO', right_on='Correo', how='left', indicator='_merge_MKT')
df_combinado = pd.merge(df_combinado, df_CP[columnas_CP], on='CORREO', how='left', indicator='_merge_CP')

# Agregar la columna 'Frecuencia'
df_combinado['Frecuencia'] = df_combinado.groupby('CORREO')['CORREO'].transform('count')

# Obtener las columnas de tipo string justo antes de aplicar la operación strip
string_columns = df_combinado.select_dtypes(include=['object']).columns

# Limpiar espacios vacíos en columnas de tipo string
for column in string_columns:
    df_combinado[column] = df_combinado[column].apply(lambda x: x.strip() if isinstance(x, str) else x)

# Eliminar filas duplicadas
df_combinado = df_combinado.drop_duplicates()

# Rellenar valores nulos con NaN
df_combinado = df_combinado.fillna(np.nan)

# Convertir las columnas de fecha a tipo datetime si no lo están
df_combinado['FECHA DE CREACIÓn'] = pd.to_datetime(df_combinado['FECHA DE CREACIÓn'])
df_combinado['Fecha-MKT'] = pd.to_datetime(df_combinado['Fecha-MKT'])
df_combinado['CREADO-CP'] = pd.to_datetime(df_combinado['CREADO-CP'])

# Crear una columna adicional con la fecha más reciente
df_combinado['FECHA MÁS RECIENTE'] = df_combinado[['FECHA DE CREACIÓn', 'Fecha-MKT', 'CREADO-CP']].max(axis=1)

# Eliminar duplicados basándote en el correo electrónico y manteniendo la fila con la fecha más reciente
df_combinado = df_combinado.sort_values(by='FECHA MÁS RECIENTE', ascending=False).drop_duplicates(subset='CORREO', keep='first')

# Guardar el DataFrame resultante en el archivo Excel, agregando a una nueva hoja
with pd.ExcelWriter(ruta_resultado, mode='a', engine='openpyxl') as writer:
    df_combinado.to_excel(writer, sheet_name='Correo_Unico_Mas_Reciente', index=False)
