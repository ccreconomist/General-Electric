import pandas as pd
import numpy as np
import warnings
import os
from pathlib import Path
from datetime import datetime, timedelta

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

ruta_bd_leads = os.path.join(carpeta_destino, "2.BD-Leads.xlsx")
ruta_Leads_MKT = os.path.join(carpeta_destino, "1.Leads_MKT.xlsx")
ruta_CP = os.path.join(carpeta_destino, "3.Clientes Potenciales.xlsx")
ruta_ventas = os.path.join(carpeta_destino, "4.Ventas.xlsx")
ruta_resultado_cp = os.path.join(carpeta_destino, "3.Clientes Potenciales.xlsx")

# Leer el archivo Excel de Clientes Potenciales
df_CP = pd.read_excel(ruta_CP, engine='openpyxl')

# Convertir la columna 'CREADO-CP' a formato datetime
df_CP['CREADO-CP'] = pd.to_datetime(df_CP['CREADO-CP'], errors='coerce')

# Ordenar por 'CREADO-CP' y mantener solo los registros más recientes por 'CLIENTE POTENCIAL'
df_CP.sort_values('CREADO-CP', ascending=False, inplace=True)
df_CP_unique = df_CP.drop_duplicates(subset='Cliente potencial', keep='first')

# Guardar el resultado en el mismo archivo Excel de Clientes Potenciales
with pd.ExcelWriter(ruta_resultado_cp, engine='openpyxl', mode='w') as writer:
    df_CP_unique.to_excel(writer, index=False)


