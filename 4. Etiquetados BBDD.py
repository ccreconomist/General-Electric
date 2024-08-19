import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

# Ruta del archivo de resultado
ruta_resultado = os.path.join(carpeta_destino, "5.Concentrado Leads.xlsx")

# Leer la hoja específica del archivo Excel
hoja_excel = "Correo_Unico_Mas_Reciente"
df = pd.read_excel(ruta_resultado, sheet_name=hoja_excel)

# Imprimir las columnas disponibles en el DataFrame
print("Columnas en el DataFrame:")
print(df.columns)

# Asegurarse de que la columna 'CANAL DE PROCEDENCIA_x' existe
if 'CANAL DE PROCEDENCIA_x' not in df.columns:
    raise KeyError("La columna 'CANAL DE PROCEDENCIA_x' no se encuentra en el DataFrame. Verifique el nombre de la columna.")

# Crear la columna de etiquetado para 'CANAL DE PROCEDENCIA_x'
def etiquetar_canal(canal):
    if pd.isna(canal):
        return None
    canal = canal.lower()
    if 'facebook' in canal or 'instagram' in canal:
        return 'Leads IG/FB'
    elif 'web' in canal or 'google' in canal:
        return 'Leads Web - Google'
    elif 'whatsapp' in canal:
        return 'Leads Whatsapp'
    elif 'sms' in canal:
        return 'Leads SMS'
    elif 'prosp. tel' in canal or 'prosp. campo' in canal:
        return 'COMERCIAL'
    else:
        return None

df['Etiqueta CANAL DE PROCEDENCIA'] = df['CANAL DE PROCEDENCIA_x'].apply(etiquetar_canal)

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

# Crear las columnas categorizadas para los días
df['Categoría DÍAS DESDE CREACIÓN_x'] = df['DÍAS DESDE CREACIÓN_x'].apply(categorizar_dias)
df['Categoría DÍAS DESDE CREACIÓN_y'] = df['DÍAS DESDE CREACIÓN_y'].apply(categorizar_dias)

# Guardar el DataFrame modificado en un nuevo archivo Excel
ruta_guardar = os.path.join(carpeta_destino, "5.Concentrado Leads_Modificado.xlsx")
df.to_excel(ruta_guardar, index=False)

print(f"Archivo modificado guardado en: {ruta_guardar}")

