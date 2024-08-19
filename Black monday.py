import pandas as pd
import warnings
import os

# Desactivar advertencias de deprecación de PyArrow
warnings.simplefilter(action='ignore', category=FutureWarning)
# Nueva ruta de la carpeta de destino
carpeta_destino = r"C:\Users\KPI_38C50\Desktop\Black Monday"
ruta_DJI = os.path.join(carpeta_destino, "data_^DJI.csv")
ruta_FCHI = os.path.join(carpeta_destino, "data_^FCHI.csv")
ruta_FTSE = os.path.join(carpeta_destino, "data_^FTSE.csv")
ruta_GDAXI = os.path.join(carpeta_destino, "data_^GDAXI.csv")
ruta_GSPC = os.path.join(carpeta_destino, "data_^GSPC.csv")
ruta_HSI = os.path.join(carpeta_destino, "data_^HSI.csv")
ruta_FCHI = os.path.join(carpeta_destino, "data_^FCHI.csv")
ruta_FCHI = os.path.join(carpeta_destino, "data_^FCHI.csv")
ruta_FCHI = os.path.join(carpeta_destino, "data_^FCHI.csv")