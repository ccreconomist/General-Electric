import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\KBD"

# Ruta del archivo de resultado
ruta_resultado = os.path.join(carpeta_destino, "2.BD-Leads.xlsx")

# Leer la hoja específica del archivo Excel
hoja_excel = "Sheet1"
df = pd.read_excel(ruta_resultado, sheet_name=hoja_excel)

# Imprimir las columnas disponibles en el DataFrame
print("Columnas en el DataFrame:")
print(df.columns)

# Verificar si la columna de días existe
columna_dias = 'Dias'  # Reemplaza 'dias' con el nombre correcto de la columna que contiene los días
if columna_dias in df.columns:
    print(f"La columna '{columna_dias}' existe en el DataFrame.")
else:
    raise ValueError(f"La columna '{columna_dias}' no se encuentra en el DataFrame.")

# Función para categorizar los días
def categorizar_dias(dias):
    if dias <= 30:
        return '0 a 30 Dias'
    elif 31 <= dias <= 60:
        return '31 a 60 Dias'
    elif 61 <= dias <= 90:
        return '61 a 90 Dias'
    elif 91 <= dias <= 120:
        return '91 a 120 Dias'
    elif 121 <= dias <= 150:
        return '121 a 150 Dias'
    elif 151 <= dias <= 180:
        return '151 a 180 Dias'
    elif dias > 181:
        return 'mayor a 181 Dias'
    else:
        return None

# Aplicar la función a la columna de días y crear una nueva columna de categorías
df['Categoria Dias'] = df[columna_dias].apply(categorizar_dias)

# Imprimir las primeras filas del DataFrame para verificar el resultado
print("Primeras filas del DataFrame con la nueva columna de categorías:")
print(df.head())

# Guardar el DataFrame modificado en un nuevo archivo Excel
ruta_resultado_modificado = os.path.join(carpeta_destino, "2.BD-Leads-Categorizado.xlsx")
df.to_excel(ruta_resultado_modificado, index=False)
