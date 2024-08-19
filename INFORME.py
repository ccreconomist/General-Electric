import pandas as pd
import warnings
import os
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

ruta_resultado = os.path.join(carpeta_destino, "Tabla 1.xlsx")

df_resultado = pd.read_excel(ruta_resultado, engine='openpyxl')
columnas_resultado = ['ESTADOS','Leads','TOTAL','CP','% Cliente Potencial','Total CP','% tasa perdida','% Op Perdidas','Conversion Ganada']





"""columnas_resultado = ['MES',	'AÑO',	'ORIGEN',	'CP',	'ASESOR',	'CORREO',	'NOMBRE',	'CELULAR_x',	'ESTATUS-LEADS',	'Estatus',	'Tipo',	'ESTADO_x',	'CIUDAD',	'CAMPAÑA',	'CANAL DE PROCEDENCIA_x',	'FECHA DE CREACIÓn',	'fecha de contacto',	'DIAS FDC',	'DIAS UTC',	'Agregado',	'contactado',	'ESPECIALIDAD_x',	'_merge_BD',	'Correo_x',	'Nombre',	'Celular',	'Canal de Procedencia',	'Fecha-MKT',	'Estado',	'Especialidad',	'COMENTARIOS',	'_merge_MKT',	'FOLIO',	'Cliente potencial',	'ULTIMO CONTACTO',	'CELULAR_y',	'ESPECIALIDAD_y',	'CANAL DE PROCEDENCIA_y',	'EVENTO/CAMPAÑA',	'ZONA',	'ESTADO_y',	'CREADO',	'Actualizado',	'Encontrado_seguimientos',	'Asesor',	'Creado',	'Resultado de actividad',	'Oportunidad',	'Tipo de seguimiento',	'Fecha',	'Estatus-Oportunidad',	'Resultado de actividad_oportunidad',	'Creado por',	'Actualizado por',	'Creado_oportunidad',	'_merge_CP',	'Frecuencia',	'Correo_y',	'Equipo',	'Condicion',	'FECHA',	'Es Cliente Ventas',	'Base-EQUIPO INSTALADO','C.P']


# Paso 1: Filtrar el DataFrame
df_frecuencia_1 = df_resultado[df_resultado['Frecuencia'] == 1].copy()

# Paso 2: Contar leads únicos basados en el correo
leads_unicos_por_correo = df_frecuencia_1['CORREO'].nunique()

# Paso 3: Clasificar los datos por año y mes
df_frecuencia_1.loc[:, 'Fecha de Creación'] = pd.to_datetime(df_frecuencia_1['FECHA DE CREACIÓn'])
df_frecuencia_1.loc[:, 'Año'] = df_frecuencia_1['Fecha de Creación'].dt.year
df_frecuencia_1.loc[:, 'Mes'] = df_frecuencia_1['Fecha de Creación'].dt.month


# Paso 4: Graficar la serie de tiempo
import matplotlib.pyplot as plt

# Calcular la cantidad de leads por mes
leads_por_mes = df_frecuencia_1.groupby(['Año', 'Mes']).size()

# Calcular la cantidad total de leads por mes y por año
leads_por_mes = df_frecuencia_1.groupby(['Año', 'Mes']).size()
leads_por_año = df_frecuencia_1.groupby('Año').size()

# Calcular el promedio de leads por mes
promedio_leads_por_mes = leads_por_mes.mean()

# Calcular el promedio de leads por año
promedio_leads_por_año = leads_por_año.mean()


print("Cantidad de Leads por Mes:")
print(leads_por_mes)
print("\nCantidad de Leads por Año:")
print(leads_por_año)
print("\nPromedio de Leads por Mes:", promedio_leads_por_mes)
print("Promedio de Leads por Año:", promedio_leads_por_año)

# ACF para ver la estacionalidad de leads por mes
from statsmodels.graphics.tsaplots import plot_acf

plot_acf(leads_por_mes)
plt.title('Autocorrelación de Leads por Mes')
plt.xlabel('Lag')
plt.ylabel('ACF')
plt.show()"""


