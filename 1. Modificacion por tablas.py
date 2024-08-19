import pandas as pd
import numpy as np
import warnings
import os
from pathlib import Path
from datetime import datetime, timedelta
import unicodedata

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\PS"

ruta_bd_leads = os.path.join(carpeta_destino, "2.BD-Leads.xlsx")
ruta_Leads_MKT = os.path.join(carpeta_destino, "1.Leads_MKT.xlsx")
ruta_CP = os.path.join(carpeta_destino, "3.Clientes Potenciales.xlsx")

# Leer los archivos Excel
df_bd_leads = pd.read_excel(ruta_bd_leads, engine='openpyxl')
df_Leads_MKT = pd.read_excel(ruta_Leads_MKT, engine='openpyxl')
df_CP = pd.read_excel(ruta_CP, engine='openpyxl')

#1-TABLA DE LEADS
# Normalizar nombres de columnas eliminando acentos
df_bd_leads.columns = [''.join((c for c in unicodedata.normalize('NFD', col) if unicodedata.category(c) != 'Mn')) for col in df_bd_leads.columns]

# Ordenar por el correo de A a Z
df_bd_leads = df_bd_leads.sort_values(by='CORREO')

# Función para limpiar y normalizar nombres
def limpiar_nombre(nombre):
    if isinstance(nombre, str):
        # Quitar acentos
        nombre = ''.join(
            (c for c in unicodedata.normalize('NFD', nombre)
             if unicodedata.category(c) != 'Mn')
        )
        # Eliminar comas
        nombre = nombre.replace(',', '')
        # Quitar dobles espacios
        nombre = ' '.join(nombre.split())
        # Poner en mayúscula la primera letra y el resto en minúsculas
        nombre = nombre.title()
    return nombre

# Aplicar la función de limpieza a la columna de Nombres
df_bd_leads['NOMBRE'] = df_bd_leads['NOMBRE'].apply(limpiar_nombre)

# Crear una columna nueva llamada 'FRECUENCIA' que cuente las veces que aparece cada correo
df_bd_leads['FRECUENCIA'] = df_bd_leads.groupby('CORREO')['CORREO'].transform('size')

# Lista de columnas a pivotar (sin columnas que causan problemas)
columnas_a_pivotar = [
    'FRECUENCIA', 'FOLIO',	'ESTATUS',	'NOMBRE',	'FECHA DE CONTACTO',	'COICIDENCIA CLIENTE POTENCIAL',	'LEAD ASOCIADO',	'ZONA',	'CELULAR',	'CORREO',	'CANAL DE PROCEDENCIA',	'TIPO DE CAMPAÑA',	'ORIGEN PROSPECTO',	'ASESOR',	'CLIENTE POTENCIAL',
    'FECHA DE ASIGNACIÓN',	'OPORTUNIDADES',	'ESPECIALIDAD',	'CIUDAD',	'ESTADO',	'CAMPAÑA',
    'COMENTARIOS',	'FECHA DE CREACIÓn',	'ULTIMO CONTACTO']

# Verificar columnas que existen en el DataFrame
columnas_existentes = [col for col in columnas_a_pivotar if col in df_bd_leads.columns]

# Función para convertir todos los valores a cadenas y unirlos con comas
def unir_valores(x):
    return ', '.join(sorted(set(str(v) for v in x.dropna())))

# Unir filas duplicadas por correo electrónico
df_unido = df_bd_leads.groupby('CORREO').agg({col: unir_valores for col in columnas_existentes}).reset_index()

# Guardar el DataFrame resultante en el mismo archivo Excel de origen
with pd.ExcelWriter(ruta_bd_leads, engine='openpyxl', mode='w') as writer:
    df_unido.to_excel(writer, index=False, sheet_name='Sheet1')

print("Archivo procesado y guardado exitosamente en el archivo de origen:", ruta_bd_leads)

#2-TABLA DE CLIENTE POTENCIAL
# Normalizar nombres de columnas eliminando acentos
df_CP.columns = [''.join((c for c in unicodedata.normalize('NFD', col) if unicodedata.category(c) != 'Mn')) for col in df_CP.columns]

# Ordenar por el correo de A a Z
df_CP = df_CP.sort_values(by='CORREO')

# Aplicar la función de limpieza a la columna de Nombres
df_CP['Cliente potencial'] = df_CP['Cliente potencial'].apply(limpiar_nombre)

# Crear una columna nueva llamada 'FRECUENCIA' que cuente las veces que aparece cada correo
df_CP['FRECUENCIA'] = df_CP.groupby('CORREO')['CORREO'].transform('size')

# Lista de columnas a pivotar (sin columnas que causan problemas)
columnas_a_pivotar_cp = [
    'FRECUENCIA', 'FOLIO', 'Cliente potencial', 'ULTIMO CONTACTO', 'CORREO', 'CELULAR', 'ESPECIALIDAD', 'CANAL DE PROCEDENCIA',
    'EVENTO/CAMPAÑA', 'ZONA', 'ESTADO', 'CIUDAD', 'ASESOR', 'CREADO-CP'
]

# Verificar columnas que existen en el DataFrame
columnas_existentes_cp = [col for col in columnas_a_pivotar_cp if col in df_CP.columns]

# Unir filas duplicadas por correo electrónico
df_unido_cp = df_CP.groupby('CORREO').agg({col: unir_valores for col in columnas_existentes_cp}).reset_index()

# Guardar el DataFrame resultante en el mismo archivo Excel de origen
with pd.ExcelWriter(ruta_CP, engine='openpyxl', mode='w') as writer:
    df_unido_cp.to_excel(writer, index=False, sheet_name='Sheet1')

print("Archivo procesado y guardado exitosamente en el archivo de origen:", ruta_CP)

#3-TABLA DE LEADS MKT
# Normalizar nombres de columnas eliminando acentos
df_Leads_MKT.columns = [''.join((c for c in unicodedata.normalize('NFD', col) if unicodedata.category(c) != 'Mn')) for col in df_Leads_MKT.columns]

# Ordenar por el correo de A a Z
df_Leads_MKT = df_Leads_MKT.sort_values(by='Correo')

# Aplicar la función de limpieza a la columna de Nombres
df_Leads_MKT['Nombre'] = df_Leads_MKT['Nombre'].apply(limpiar_nombre)

# Crear una columna nueva llamada 'FRECUENCIA' que cuente las veces que aparece cada correo
df_Leads_MKT['FRECUENCIA'] = df_Leads_MKT.groupby('Correo')['Correo'].transform('size')

# Lista de columnas a pivotar
columnas_a_pivotar_mkt = [
    'Procedencia', 'Fecha-MKT', 'Nombre', 'Especialidad', 'Correo', 'Celular',
    'Ciudad', 'Estado', 'Campaña', 'NOTAS'
]

# Verificar columnas que existen en el DataFrame
columnas_existentes_mkt = [col for col in columnas_a_pivotar_mkt if col in df_Leads_MKT.columns]

# Unir filas duplicadas por correo electrónico
df_unido_mkt = df_Leads_MKT.groupby('Correo').agg({col: unir_valores for col in columnas_existentes_mkt}).reset_index()

# Guardar el DataFrame resultante en el mismo archivo Excel de origen
with pd.ExcelWriter(ruta_Leads_MKT, engine='openpyxl', mode='w') as writer:
    df_unido_mkt.to_excel(writer, index=False, sheet_name='Sheet1')

print("Archivo procesado y guardado exitosamente en el archivo de origen:", ruta_Leads_MKT)
