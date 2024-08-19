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
ruta_BI_13_14 = os.path.join(carpeta_destino, "BI_13_14.xlsx")
ruta_BI_14_24 = os.path.join(carpeta_destino, "BI_14_24.xlsx")
ruta_BI_23_24 = os.path.join(carpeta_destino, "BI_23_24.xlsx")
ruta_BI_10_21 = os.path.join(carpeta_destino, "BI-10_21.xlsx")
ruta_BI_R = os.path.join(carpeta_destino, "Base Instalada a Julio 2024.xlsx")


# Seleccionar columnas
columnas_BI_13_14 = ['No. Pedido Kpi','GON.','Vendedor','Cliente ','Direccion','CORREO','Telefonos','Fecha Max. de Entrega',' GE entrego (fecha)',' KPI entrego (fecha)','Marca-Modelo','NUM. SER.','Version de Software ','Trans. Convexo','Num. Serie','Trans. Convexo Vol.','Num. Serie','Trans. Endo-cavitario','Num. Serie','Trans. Endo-cavitario Vol.','Num. Serie','Trans. Lineal','Num. Serie','Otro','Perifericos','Requerimientos Especiales','Opciones abiertas','Clave','INGENIERO','Equipo a Prestamo','Nota','Llamada de servicio - pendiente de entrega al cliente','Forma de pago','Fecha del Anticipo']
columnas_BI_14_24 = ['Cliente ','Fecha de salida','Fecha Entrega cliente','Fecha de instalacion','ing','Marca-Modelo','Version','Condicion','GON.','SID','Asesor','Estado','Ciudad','Columna1','CP','ESTADO_2','Años']
columnas_BI_23_24 = ['TELEFONO','CLIENTE','Val','TRANSPORTO','MES','FECHA DE SALIDA','FECHA ENTREGA AL CLIENTE','FECHA DE INSTALACION','INSTALO','MARCA-MODELO','VERSION','CONDICION','GON','NUM. SER.','ASESOR','ESTADO','CIUDAD','DIRECCION']
columnas_BI_10_21 = ['NUMERO DE CLIENTE','COMENTARIOS','CLIENTE','FECHA DE LA VENTA','FECHA DE NACIMIENTO','TELEFONO 1','TELEFONO 2','CIUDAD','ESTADO','DIRECCION','RFC','SOCIO','CLINICA / HOSPITAL','ESPECIALIDAD','AÑO DE INSTALACION','MARCA','MODELO','VERSION','PRECIO','MONEDA','NO. DE SERIE','CONDICIÓN','GARANTIA','ASESOR','FORMA DE COMPRA CONTADO/FINANCIADO','PLAZO','FINALIZA','N° SERVICIOS']
columnas_BI_R = ['ID', 'Cliente', 'Correo', 'Metodo de pago', 'fechaContrato', 'fechaTermino', 'Estatus', 'Tipo', 'Plazos', 'Fecha de salida', 'Fecha Entrega cliente', 'Fecha de instalacion', 'ing', 'Marca-Modelo', 'Version', 'Condicion', 'GON.', 'SID', 'Asesor', 'Estado', 'Ciudad', 'Columna1', 'CP', 'ESTADO_2', 'Años', 'Especialidad', 'Canal', 'Clase_Minima', 'Clase_Maxima', 'Clase_Predominante']

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

df_BI_13_14 = leer_excel(ruta_BI_13_14, columnas_BI_13_14)
df_BI_14_24 = leer_excel(ruta_BI_14_24, columnas_BI_14_24)
df_BI_23_24 = leer_excel(ruta_BI_23_24, columnas_BI_23_24)
df_BI_10_21 = leer_excel(ruta_BI_10_21, columnas_BI_10_21)
df_BI_R = leer_excel(ruta_BI_R, columnas_BI_R)


# Agregar una columna 'Origen' para identificar la tabla original
df_BI_13_14['Origen'] = 'BI_13_14'
df_BI_14_24['Origen'] = 'BI_14_24'
df_BI_23_24['Origen'] = 'BI_23_24'
df_BI_10_21['Origen'] = 'BI_10_21'
df_BI_R['Origen'] = 'BI_R'


# Renombrar las columnas 'Cliente' para uniformidad
df_BI_13_14.rename(columns={'Cliente ': 'Cliente'}, inplace=True)
df_BI_14_24.rename(columns={'Cliente ': 'Cliente'}, inplace=True)
df_BI_23_24.rename(columns={'CLIENTE': 'Cliente'}, inplace=True)
df_BI_10_21.rename(columns={'CLIENTE': 'Cliente'}, inplace=True)
df_BI_R.rename(columns={'Cliente': 'Cliente'}, inplace=True)


# Concatenar todos los DataFrames
df_todos = pd.concat([df_BI_13_14, df_BI_14_24, df_BI_23_24, df_BI_10_21, df_BI_R], ignore_index=True)

# Crear un identificador único para cada cliente
if 'Cliente' in df_todos.columns:
    df_todos['ID_Cliente'] = df_todos.groupby('Cliente').ngroup()
else:
    raise ValueError("La columna 'Cliente' no se encuentra en el DataFrame concatenado.")

# Pivotar los datos y mantener la columna 'Origen'
df_pivot = df_todos.pivot_table(index=['ID_Cliente', 'Cliente'], columns='Origen', aggfunc='first')

# Ajustar la tabla pivote para que 'Origen' sea una columna
df_pivot = df_pivot.reset_index()
df_pivot.columns = [f'{col[0]}_{col[1]}' if col[1] else col[0] for col in df_pivot.columns]

# Guardar el resultado en un nuevo archivo Excel
ruta_salida = os.path.join(carpeta_destino, "BASE_INSTALADA_ACTUALIZADA.xlsx")
df_pivot.to_excel(ruta_salida, index=False)

print(f"Archivo guardado en: {ruta_salida}")

