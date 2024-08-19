import pandas as pd
import warnings
import os
from pathlib import Path

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\KBD"

ruta_seguimientos = os.path.join(carpeta_destino, "7.Seguimientos.xlsx")
ruta_oportunidad = os.path.join(carpeta_destino, "8.Seguimiento OP.xlsx")
ruta_historial = os.path.join(carpeta_destino, "12. Historial.xlsx")

# Leer los archivos Excel
df_oportunidades = pd.read_excel(ruta_oportunidades, engine='openpyxl')
df_seguimientos = pd.read_excel(ruta_seguimientos, engine='openpyxl')
df_oportunidad = pd.read_excel(ruta_oportunidad, engine='openpyxl')
df_historial = pd.read_excel(ruta_historial, engine='openpyxl')

# Definir columnas a seleccionar

# Seguimientos
columnas_seguimientos = ['Folio', 'Cliente potencial', 'Asesor2', 'Tipo', "fecha de contacto-Agenda", 'Estatus', 'Fecha Real', 'Resultado de actividad']
# Seguimientos oportunidad
columnas_oportunidad = ['Cliente potencial', 'Asesor3', 'Estatus de oportunidad', 'Concepto Venta', 'Transductores y/o SW', 'Celular', 'Teléfono', 'Estado', 'Origen de oportunidad', 'Congreso y evento', 'Campaña de Facebook', 'Titulo', 'Folio', 'Monto', '% Conversión', 'Fecha Prob.Cierre', 'Atendido por el asesor', 'Tipo de campaña FB', 'Campaña Comercial', 'Campaña Ongoing', 'Tipo de campaña Google', 'Motivos de rechazo', 'Campaña de Google', 'Creado por', 'Actualizado por', 'Creado-OP', 'Actualizado1']
# Historial
columnas_historial = ['Folio-h', 'Cliente potencial h', 'Asesor actual h', 'Asesor previo h', 'Actualizado por h', 'Creado h']

# Seleccionar las columnas necesarias de cada DataFrame
df_seguimientos = df_seguimientos[columnas_seguimientos]
df_oportunidad = df_oportunidad[columnas_oportunidad]
df_historial = df_historial[columnas_historial]

#Necesito que abras el archivo de seguimiento y seguimiento OP, donde