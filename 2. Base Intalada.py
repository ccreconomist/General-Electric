import pandas as pd
import numpy as np
import unicodedata
import os
import warnings

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD\PS"
ruta_Base_Actualizada = os.path.join(carpeta_destino, "BASE_INSTALADA_ACTUALIZADA.xlsx")

# Función para eliminar acentos
def quitar_acentos(texto):
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    return texto

# Cargar el archivo Excel
df = pd.read_excel(ruta_Base_Actualizada)

# Revisar si la columna "Cliente" existe
if 'Cliente' in df.columns:
    # Imprimir los valores originales
    print("Valores originales de la columna 'Cliente':")
    print(df['Cliente'].head())

    # Modificar los datos de la columna "Cliente"
    df['Cliente'] = df['Cliente'].astype(str)  # Asegurarse de que todos los valores sean cadenas
    df['Cliente'] = df['Cliente'].str.strip()  # Eliminar espacios al inicio y al final
    df['Cliente'] = df['Cliente'].str.replace(r'\s+', ' ', regex=True)  # Eliminar dobles espacios
    df['Cliente'] = df['Cliente'].apply(quitar_acentos)  # Eliminar acentos
    df['Cliente'] = df['Cliente'].str.title()  # Convertir a formato de título

    # Imprimir los valores modificados
    print("\nValores modificados de la columna 'Cliente':")
    print(df['Cliente'].head())

    # Guardar los cambios en el archivo original
    df.to_excel(ruta_Base_Actualizada, index=False)
else:
    print("La columna 'Cliente' no se encuentra en el DataFrame.")




