import pandas as pd
import os

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\KBD"

# Rutas de los archivos
ruta_leads = os.path.join(carpeta_destino, "Leads.xlsx")
ruta_Clientes_Potenciales = os.path.join(carpeta_destino, "Clientes_Potenciales.xlsx")

# Leer los archivos Excel
df_leads = pd.read_excel(ruta_leads)
df_clientes_potenciales = pd.read_excel(ruta_Clientes_Potenciales)

# Asegurarnos de que ambas columnas de correo tienen el mismo nombre
df_leads.rename(columns={'CORREO': 'CORREO ELÉCTRONICO'}, inplace=True)

# Identificar duplicados en la columna de correo en cada DataFrame
df_leads['DUPLICADO_EN_LEADS'] = df_leads.duplicated(subset=['CORREO ELÉCTRONICO'], keep=False)
df_clientes_potenciales['DUPLICADO_EN_CLIENTES'] = df_clientes_potenciales.duplicated(subset=['CORREO ELÉCTRONICO'], keep=False)

# Unir los DataFrames por la columna de correo
df_merged = pd.merge(df_leads, df_clientes_potenciales, on='CORREO ELÉCTRONICO', how='inner', suffixes=('_LEADS', '_CLIENTES'))

# Mostrar las primeras filas del DataFrame unido con las columnas de duplicados
print(df_merged.head())

# Guardar el DataFrame unido con los duplicados marcados en un nuevo archivo Excel
df_merged.to_excel(os.path.join(carpeta_destino, "Leads_Clientes_Potenciales_Unidos.xlsx"), index=False)
