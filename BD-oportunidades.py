import pandas as pd
import os
import warnings

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

# Rutas de los archivos Excel
ruta_Oportunidades = os.path.join(carpeta_destino, "6.Oportunidades.xlsx") # BASE DE OPORTUNIDADES

# Leer los archivos Excel
df_Oportunidades = pd.read_excel(ruta_Oportunidades, engine='openpyxl')

# Paso 1: Obtener clientes potenciales únicos y contar las veces que aparecen con cada "Estatus de oportunidad"
df_contador = df_Oportunidades.groupby(['Cliente potencial', 'Estatus de oportunidad']).size().unstack(fill_value=0)

# Paso 2: Obtener la fecha más reciente para cada "Cliente Potencial" y su correspondiente "Estatus de oportunidad"
"""df_fecha_max = df_Oportunidades.groupby(['Cliente potencial', 'Estatus de oportunidad'])[['Creado', 'Actualizado']].max()

# Paso 3: Unir las dos tablas anteriores para obtener la columna "Concepto de venta" más reciente
df_resultado = pd.merge(df_fecha_max, df_contador, left_index=True, right_index=True)

# Convertir las columnas relevantes a tipo numérico
df_resultado[['Perdida', 'Ganada', 'Cotizacion', 'Negociacion', 'Analisis', 'Demostracion', 'Pre-Cierre', 'Venta Cerrada']] = df_resultado[['Perdida', 'Ganada', 'Cotizacion', 'Negociacion', 'Analisis', 'Demostracion', 'Pre-Cierre', 'Venta Cerrada']].apply(pd.to_numeric, errors='coerce')

# Paso 4: Encontrar el "Estatus de oportunidad" con la mayor frecuencia para cada "Cliente Potencial"
df_resultado['Estatus_mas_frecuente'] = df_resultado.idxmax(axis=1)

# Paso 5: Seleccionar la columna "Concepto de venta" correspondiente al "Estatus_mas_frecuente"
df_resultado['Concepto_de_venta_mas_reciente'] = df_resultado.apply(
    lambda row: df_Oportunidades.loc[
        (df_Oportunidades['Cliente potencial'] == row.name[0]) &
        (df_Oportunidades['Estatus de oportunidad'] == row['Estatus_mas_frecuente']),
        'Concepto de ventas'].max(), axis=1
)

# Resetear el índice para tener "Cliente Potencial" como una columna
df_resultado.reset_index(inplace=True)

# Guardar el DataFrame modificado en un nuevo archivo Excel
ruta_resultado = r"C:\Users\KPI_38C50\Desktop\BD\Oportunidades_Modificado.xlsx"
df_resultado.to_excel(ruta_resultado, index=False, engine='openpyxl')
