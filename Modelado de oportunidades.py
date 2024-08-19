import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

# Ruta del archivo Excel
ruta_OP = os.path.join(carpeta_destino, "resultado_oportunidades.xlsx")

# Leer el archivo Excel
df_OP = pd.read_excel(ruta_OP, engine='openpyxl')

# Columnas especificadas
columnas_OP = ['Año', 'Mes', 'Celular', 'Teléfono', 'Canal', 'Origen Op', 'Campaña', 'Cliente potencial', 'Estatus', 'Asesor', 'Concepto Venta', 'Estado', 'Origen de oportunidad', 'Congreso y evento', 'Campaña de Facebook', 'Titulo', 'Folio', 'Monto', '% Conversión', 'Fecha Prob.Cierre', 'Tipo de campaña FB', 'Campaña Comercial', 'Tipo de campaña Google', 'Motivos de rechazo', 'Campaña de Google', 'Creado por', 'Actualizado por', 'Fecha creacion de Cliente potencial', 'CREADO.CP', 'Actualizado', 'Dias transcurridos', 'DFDC', 'Dias transcurridos entre agregado y oportunidad']

# Filtrar columnas que existen en el archivo
columnas_presentes = [col for col in columnas_OP if col in df_OP.columns]

# Seleccionar solo las columnas presentes
df_OP = df_OP[columnas_presentes]


"""# Contar clientes únicos por estatus de oportunidad y asesor
resumen_clientes = df_OP.groupby(['Estatus', 'Asesor'])['Cliente potencial'].nunique().reset_index()
resumen_clientes.rename(columns={'Cliente potencial': 'Clientes únicos'}, inplace=True)

# Calcular el promedio de días en cada estatus de oportunidad
resumen_dias = df_OP.groupby('Estatus')['Dias transcurridos entre agregado y oportunidad'].mean().reset_index()
resumen_dias.rename(columns={'Dias transcurridos entre agregado y oportunidad': 'Días promedio'}, inplace=True)

# Unir los dos resúmenes
resumen_final = pd.merge(resumen_clientes, resumen_dias, on='Estatus', how='left')

# Guardar el resumen en un archivo Excel
ruta_resumen = os.path.join(carpeta_destino, "resumen_oportunidades.xlsx")
resumen_final.to_excel(ruta_resumen, index=False)

print(columnas_presentes)"""




