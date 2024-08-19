import pandas as pd
import warnings
import os
from pathlib import Path

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

ruta_oportunidades = os.path.join(carpeta_destino, "3.Clientes Potenciales.xlsx")
ruta_seguimientos = os.path.join(carpeta_destino, "7.Seguimientos.xlsx")
ruta_oportunidad = os.path.join(carpeta_destino, "8.Seguimiento OP.xlsx")
ruta_OP = os.path.join(carpeta_destino, "6.Oportunidades.xlsx")


# Leer los archivos Excel
df_oportunidades = pd.read_excel(ruta_oportunidades, engine='openpyxl')
df_seguimientos = pd.read_excel(ruta_seguimientos, engine='openpyxl')
df_oportunidad = pd.read_excel(ruta_oportunidad, engine='openpyxl')
df_OP = pd.read_excel(ruta_OP, engine='openpyxl')

# Definir columnas a seleccionar
columnas_oportunidades = ['FOLIO','Cliente potencial','ULTIMO CONTACTO','CORREO','CELULAR','ESPECIALIDAD','CANAL DE PROCEDENCIA','EVENTO/CAMPAÑA','ZONA','ESTADO','CIUDAD','ASESOR','CREADO-CP']
columnas_seguimientos = ['Cliente potencial', 'Asesor', 'Tipo', 'fecha de contacto-Agenda', 'Estatus', 'Fecha Real', 'Resultado de actividad']
columnas_oportunidad = ['Folio','Oportunidad','Cliente potencial','Tipo de seguimiento','Fecha','Estatus','Resultado de actividad','Creado por','Actualizado por','Creado-OP','Actualizado1']
columnas_OP = ['Año','Mes','Cliente potencial','Asesor','Estatus de oportunidad','Concepto Venta','Transductores y/o SW','Celular','Teléfono','Estado','Origen de oportunidad','Congreso y evento','Campaña de Facebook','Titulo','Folio','Monto','% Conversión','Fecha Prob.Cierre','Tipo de campaña FB','Campaña Comercial','Tipo de campaña Google','Motivos de rechazo','Campaña de Google','Creado por','Actualizado por','CREADO.CP','Actualizado']

# Seleccionar las columnas necesarias de cada DataFrame
df_oportunidades = df_oportunidades[columnas_oportunidades]
df_seguimientos = df_seguimientos[columnas_seguimientos]
df_oportunidad = df_oportunidad[columnas_oportunidad]
df_OP = df_OP[columnas_OP]

# Convertir la columna 'Actualizado1' a tipo datetime
# df_OP['Actualizado'] = pd.to_datetime(df_OP['Actualizado'])

# 1. Obtener valores únicos de 'Cliente potencial' de la tabla de seguimientos
clientes_seguimientos_unicos = df_seguimientos['Cliente potencial'].unique()
clientes_oportunidad_unicos = df_oportunidad['Cliente potencial'].unique()
clientes_oportunidades_unicos = df_oportunidades['Cliente potencial'].unique()
clientes_OP_unicos = df_OP['Cliente potencial'].unique()


# 2. Obtener la fecha de contacto más reciente por 'Cliente potencial' de la tablas
ultima_fecha_contacto = df_seguimientos.groupby('Cliente potencial')['Fecha Real'].max().reset_index()
ultima_fecha_actualizado = df_oportunidad.groupby('Cliente potencial')['Creado-OP'].max().reset_index()
ultima_fecha_actualizado1 = df_OP.groupby('Cliente potencial')['CREADO.CP'].max().reset_index()

# 3. Unir la información de seguimientos con la tabla de clientes potenciales
df_final = pd.merge(df_oportunidades, ultima_fecha_contacto.merge(ultima_fecha_actualizado.merge(ultima_fecha_actualizado1,
                                    on='Cliente potencial', how='left'), on='Cliente potencial', how='left'), on='Cliente potencial', how='left')
# 4. Marcar en la tabla de clientes potenciales cuáles clientes están presentes en los seguimientos
df_final['Encontrado_seguimientos'] = df_final['Cliente potencial'].isin(clientes_seguimientos_unicos)

# 5. Agregar las columnas relevantes de seguimientos a la tabla de clientes potenciales
df_final = pd.merge(df_final, df_seguimientos, on=['Cliente potencial', 'Fecha Real'], how='left', suffixes=('', '_seguimiento'))

# 6. Agregar las columnas relevantes de seguimientos a la tabla de clientes potenciales
df_final = pd.merge(df_final, df_oportunidad, on=['Cliente potencial', 'Creado-OP'], how='left', suffixes=('', '_oportunidad'))

# 6. Agregar las columnas relevantes de seguimientos a la tabla de clientes potenciales
df_final = pd.merge(df_final, df_OP, on=['Cliente potencial', 'CREADO.CP'], how='left', suffixes=('', '_oportunidad'))

# 7. Guardar los resultados en el mismo archivo '3.Clientes Potenciales'
df_final.to_excel(ruta_oportunidades, index=False)

print("Resultados guardados en:", ruta_oportunidades)
