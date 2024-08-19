from geopy.geocoders import Nominatim

# Lista de direcciones
# Crear un objeto geolocalizador
direcciones = [
'Luis Donaldo Colosio Murrieta 106 Lomas del Campestre II'
]
geolocalizador = Nominatim(user_agent="CcrViridiana")
# Función para obtener coordenadas de una dirección
def obtener_coordenadas(direccion):
    try:
        ubicacion = geolocalizador.geocode(direccion)
        if ubicacion:
            return ubicacion.latitude, ubicacion.longitude
    except Exception as e:
        print(f"No se pudo obtener la ubicación para {direccion}: {e}")
    return None, None
# Obtener coordenadas para todas las direcciones
coordenadas = [obtener_coordenadas(d) for d in direcciones]
# Mostrar las coordenadas obtenidas
for i, coord in enumerate(coordenadas):
    print(f"Dirección {i+1}: {coord}")

