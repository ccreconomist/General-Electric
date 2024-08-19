import pandas as pd
import warnings
import os
import matplotlib.pyplot as plt

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\PS"

ruta_Base = os.path.join(carpeta_destino, "5.Concentrado Leads.xlsx")  # BASE CONCENTRADO LEADS
ruta_Oportunidades = os.path.join(carpeta_destino, "6.Oportunidades.xlsx")  # BASE DE OPORTUNIDADES

df_Base = pd.read_excel(ruta_Base, engine='openpyxl')
df_Oportunidades = pd.read_excel(ruta_Oportunidades, engine='openpyxl')

# Paso 1: Obtener clientes potenciales únicos y contar las veces que aparecen con cada "Estatus de oportunidad"
df_contador = df_Oportunidades.groupby(['Cliente potencial', 'Estatus de oportunidad']).size().unstack(fill_value=0)

# Mapeo de Concepto Venta a números
concepto_venta_mapping =\
{
 'Voluson E10 BT19'	:	1,
 'Voluson E6 BT13'	:	2,
 'VOLUSON E6 BT21'	:	3,
 'Voluson P8 BT18 Console High Estandar'	:	4,
 'Voluson S10 Expert BT18'	:	5,
 'Voluson S8 Touch BT18'	:	6,
 'Voluson Swift+'	:	7,
 'Voluson P8 BT18 Console High'	:	8,
 'Voluson SWIFT+ BT23'	:	9,
 'LOGIQ F8 EXPERT 4D version'	:	10,
 'LOGIQ V5 EXPERT'	:	11,
 'Versana Premier VA w 4D'	:	12,
 'Equipo Usuado'	:	13,
 'Voluson E6 BT19'	:	14,
 'Voluson S8 BT18'	:	15,
 'Voluson SWIFT+ New'	:	16,
 'LOGIQ F6 R2 4D version'	:	17,
 'Voluson E8 BT21 LCD'	:	18,
 'VP8, P6 BT18 (eIFU) Paperless Manual Kit'	:	19,
 'Versana Active™ 4D Ready'	:	20,
 'Voluson E8 BT19'	:	21,
 'Voluson E10 BT21 OLED'	:	22,
 'VS10 Expert, S10 BT18 (eIFU) Paperless Manual Kit'	:	23,
 'Logiq P5 Powered By Voluson BT11'	:	24,
 'H48711FB-Voluson-E10-BT20'	:	25,
 'S8-S6 Drawer'	:	26,
 'LOGIQ F8 R2 Expert LA SW kits (easy 3D, LOGIQ View, Advance 3D, Report, Bflow/ Bflow color)'	:	27,
 'AN Key Cap Kit - VP8-P6: Spanish'	:	28,
 'Voluson E10 BT19 OLED Monitor': 29,
 'Voluson Swift+ BR'	:	30,
 'VOLUSON E6 BT18' : 31
}

# Aplicar el mapeo a la columna 'Concepto Venta'
df_Oportunidades['Concepto_Venta_Num'] = df_Oportunidades['Concepto Venta'].map(concepto_venta_mapping)

# Filtrar por rango de fechas
fecha_inicio = pd.to_datetime('2024-01-01')
fecha_fin = pd.to_datetime('2024-07-31')

# Convertir la columna 'Fecha Prob. Cierre' a datetime
df_Oportunidades['Fecha'] = pd.to_datetime(df_Oportunidades['Fecha'], errors='coerce')

# Filtrar por rango de fechas
df_filtrado_fecha = df_Oportunidades[
    (df_Oportunidades['Fecha'].notnull()) &
    (df_Oportunidades['Fecha'] >= fecha_inicio) &
    (df_Oportunidades['Fecha'] <= fecha_fin)
]

# Crear una tabla de frecuencia por fecha y estatus
tabla_frecuencia_fecha_estatus = df_filtrado_fecha.groupby(['Cliente potencial', 'Estatus de oportunidad', 'Concepto_Venta_Num']).size().unstack(fill_value=0)

# Exportar la tabla de frecuencia por fecha y estatus a Excel
ruta_resultado_fecha_estatus = os.path.join(carpeta_destino, "Frecuencia_concepto.xlsx")
tabla_frecuencia_fecha_estatus.to_excel(ruta_resultado_fecha_estatus, index=True, engine='openpyxl')
print(f"La tabla de frecuencia para el rango de fechas y estatus ha sido exportada a: {ruta_resultado_fecha_estatus}")
