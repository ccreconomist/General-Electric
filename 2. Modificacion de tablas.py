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

# 1-TABLA DE LEADS

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

# Lista de columnas a pivotar
columnas_a_pivotar = [
    'FRECUENCIA', 'TIPO DE LEAD', 'FECHA DE CONTACTO', 'FECHA DE CREACIÓn', 'ULTIMO CONTACTO',
    'COINCIDENCIA CLIENTE POTENCIAL', 'LEAD ASOCIADO', 'ZONA', 'CELULAR', 'CORREO',
    'CANAL DE PROCEDENCIA', 'TIPO DE CAMPAÑA', 'ORIGEN PROSPECTO', 'ASESOR',
    'CLIENTE POTENCIAL', 'FECHA DE ASIGNACIÓN', 'OPORTUNIDADES', 'ESPECIALIDAD',
    'CIUDAD', 'ESTADO', 'CAMPAÑA', 'COMENTARIOS']

# Función para convertir todos los valores a cadenas y unirlos con comas
def unir_valores(x):
    return ','.join(sorted(set(str(v) for v in x.dropna())))

# Unir filas duplicadas por correo electrónico
df_unido = df_bd_leads.groupby('CORREO').agg(unir_valores).reset_index()

# Guardar el DataFrame resultante en el mismo archivo Excel de origen
with pd.ExcelWriter(ruta_bd_leads, engine='openpyxl', mode='w') as writer:
    df_unido.to_excel(writer, index=False, sheet_name='Sheet1')

print("Archivo procesado y guardado exitosamente en el archivo de origen:", ruta_bd_leads)

# 2-TABLA DE CLIENTE POTENCIAL

# Asegurarse de que todos los valores en la columna 'CORREO' sean cadenas
df_CP['CORREO'] = df_CP['CORREO'].astype(str)

# Ordenar por el correo de A a Z
df_CP = df_CP.sort_values(by='CORREO')

# Función para limpiar y normalizar nombres
def limpiar_nombre_1(nombre):
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
df_CP['Cliente potencial'] = df_CP['Cliente potencial'].apply(limpiar_nombre_1)

# Crear una columna nueva llamada 'FRECUENCIA' que cuente las veces que aparece cada correo
df_CP['FRECUENCIA'] = df_CP.groupby('CORREO')['CORREO'].transform('size')

# Lista de columnas a pivotar
columnas_a_pivotar = [
    'FRECUENCIA', 'FOLIO', 'Cliente potencial', 'ULTIMO CONTACTO', 'CORREO', 'CELULAR', 'ESPECIALIDAD', 'CANAL DE PROCEDENCIA',
    'EVENTO/CAMPAÑA', 'ZONA', 'ESTADO', 'CIUDAD', 'ASESOR', 'CREADO-CP']

# Función para convertir todos los valores a cadenas y unirlos con comas
def unir_valores(x):
    return ','.join(sorted(set(str(v) for v in x.dropna())))

# Unir filas duplicadas por correo electrónico
df_unido_cp = df_CP.groupby('CORREO').agg(unir_valores).reset_index()

# Guardar el DataFrame resultante en el mismo archivo Excel de origen
with pd.ExcelWriter(ruta_CP, engine='openpyxl', mode='w') as writer:
    df_unido_cp.to_excel(writer, index=False, sheet_name='Sheet1')

print("Archivo procesado y guardado exitosamente en el archivo de origen:", ruta_CP)

# 3-TABLA DE LEADS MKT

# Verificar los nombres de las columnas
print("Columnas de df_Leads_MKT:", df_Leads_MKT.columns)

# Asegurarse de que todos los valores en la columna 'Correo' sean cadenas (ajusta el nombre de la columna si es necesario)
df_Leads_MKT['Correo'] = df_Leads_MKT['Correo'].astype(str)

# Ordenar por el correo de A a Z
df_Leads_MKT = df_Leads_MKT.sort_values(by='Correo')

# Función para limpiar y normalizar nombres
def limpiar_nombre_2(nombre):
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
df_Leads_MKT['Nombre'] = df_Leads_MKT['Nombre'].apply(limpiar_nombre_2)

# Crear una columna nueva llamada 'FRECUENCIA' que cuente las veces que aparece cada correo
df_Leads_MKT['FRECUENCIA'] = df_Leads_MKT.groupby('Correo')['Correo'].transform('size')

# Lista de columnas a pivotar
columnas_a_pivotar_mkt = [
    'Procedencia', 'Fecha-MKT', 'Nombre', 'Especialidad', 'Correo', 'Celular',
    'Ciudad', 'Estado', 'Campaña', 'NOTAS'
]

# Función para convertir todos los valores a cadenas y unirlos con comas
def unir_valores(x):
    return ','.join(sorted(set(str(v) for v in x.dropna())))

# Unir filas duplicadas por correo electrónico
df_unido_mkt = df_Leads_MKT.groupby('Correo').agg(unir_valores).reset_index()

# Guardar el DataFrame resultante en el mismo archivo Excel de origen
with pd.ExcelWriter(ruta_Leads_MKT, engine='openpyxl', mode='w') as writer:
    df_unido_mkt.to_excel(writer, index=False, sheet_name='Sheet1')

print("Archivo procesado y guardado exitosamente en el archivo de origen:", ruta_Leads_MKT)
