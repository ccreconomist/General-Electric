import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)

# Ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\BD"

# Ruta del archivo de resultado
ruta_resultado = os.path.join(carpeta_destino, "NSE.xlsx")

# Cargar el archivo de Excel
df = pd.read_excel(ruta_resultado)

# Paso 2: Agrupar los datos por estado y clase socioeconómica, y calcular la suma de la población por cada combinación
grupo_estado_clase = df.groupby(['Estado', 'Clase_Socioeconomica'])['Poblacion'].sum().reset_index()

# Paso 3: Encontrar el mínimo, máximo y predominancia por estado
resultados = grupo_estado_clase.groupby('Estado').agg(
    Clase_Minima=('Poblacion', 'min'),
    Clase_Maxima=('Poblacion', 'max'),
    Clase_Predominante=('Poblacion', lambda x: grupo_estado_clase.loc[x.idxmax(), 'Clase_Socioeconomica'])
).reset_index()

# Mostrar los resultados
print(resultados)

# Paso 4: (Opcional) Visualizar los resultados
import matplotlib.pyplot as plt

# Graficar la clase mínima por estado
plt.figure(figsize=(10, 6))
plt.bar(resultados['Estado'], resultados['Clase_Minima'], color='blue', alpha=0.7, label='Clase Mínima')
plt.bar(resultados['Estado'], resultados['Clase_Maxima'], color='red', alpha=0.7, label='Clase Máxima')
plt.xlabel('Estado')
plt.ylabel('Población')
plt.title('Clase Socioeconómica Mínima y Máxima por Estado')
plt.legend()
plt.show()

# Graficar la clase predominante por estado
plt.figure(figsize=(10, 6))
plt.bar(resultados['Estado'], resultados['Clase_Predominante'], color='green')
plt.xlabel('Estado')
plt.ylabel('Clase Predominante')
plt.title('Clase Socioeconómica Predominante por Estado')
plt.show()


#Tengo 209 registros con las siguientes variables, donde cada columna es una variable estas son las oclumnas
#['Id','Especialidad','Metodo de pago','Tipo Equipo','Equipo','Estado','Procedencia','INICIALES','fecha de nacimiento','Homoclave','años','Clase_Minima','Clase_Maxima','Clase_Predominante_xEstado','A/B',"'C+','C','C-','D+','D/E','Mínimo:','Máximo:','Promedio:']
#todas la columnas son texto menos las dos columnas de Clase_Minima, Clase_Maxima, 'A/B',"'C+','C','C-','D+','D/E','Mínimo:','Máximo:','Promedio:' es numero sin decimal
#lO QUE NECESITO QUE HAGA ES UN ANALISIS POR SEGEMENTACION
#anticipiso cuando son arrendamientos, meses, plazos
#anticipos en % cuando son financiamiento intern, y si tomo equipo o no, del 100% a cuantos les tiene que comprar su equipo
#







