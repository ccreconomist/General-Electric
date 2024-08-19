import pandas as pd
import warnings
from datetime import datetime
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

# Ruta del archivo de resultado
ruta_resultado = os.path.join(carpeta_destino, "4.Ventas.xlsx")

# Cargar el archivo de Excel
df = pd.read_excel(ruta_resultado)

# Asegurarse de que todos los valores en la columna 'Rfc' sean cadenas y reemplazar NaN por una cadena vacía
df['Rfc'] = df['Rfc'].fillna('').astype(str)

# Reemplazar espacios en blanco con 'XXXX' en la columna 'Rfc'
df['Rfc'] = df['Rfc'].str.replace(' ', 'XXXX')

# Crear la columna 'tipo' para categorizar según la longitud del Rfc
df['tipo'] = df['Rfc'].apply(lambda x: 'Persona Moral' if len(x) == 12 else 'Persona Física' if len(x) == 13 else 'Desconocido')

# Crear la columna 'INICIALES' con las primeras 4 letras del Rfc
df['INICIALES'] = df['Rfc'].str[:4]

# Crear la columna 'fecha de nacimiento' tomando de la posición 5 a la 10 del Rfc y formateando en 'aa-mm-dd'
df['fecha de nacimiento'] = df['Rfc'].str[4:10].apply(lambda x: f"{x[:2]}-{x[2:4]}-{x[4:]}")

# Crear la columna 'Homoclave' tomando de la posición 11 a 13 del Rfc
df['Homoclave'] = df['Rfc'].str[10:13]

# Calcular la columna 'años' a partir de 'fecha de nacimiento', excluyendo 'Persona Moral'
def calcular_edad(fecha_nac_str):
    try:
        fecha_nac = datetime.strptime(fecha_nac_str, '%y-%m-%d')
        hoy = datetime.today()
        if fecha_nac > hoy:  # Si la fecha de nacimiento es posterior a la fecha actual
            return None  # Devolver None en lugar de un valor negativo
        edad = hoy.year - fecha_nac.year - ((hoy.month, hoy.day) < (fecha_nac.month, fecha_nac.day))
        return max(edad, 0)  # Devolver 0 si la edad calculada es negativa
    except:
        return None

df['años'] = df.apply(lambda row: calcular_edad(row['fecha de nacimiento']) if row['tipo'] != 'Persona Moral' else None, axis=1)

# Guardar el DataFrame modificado en el archivo de Excel
df.to_excel(ruta_resultado, index=False)
