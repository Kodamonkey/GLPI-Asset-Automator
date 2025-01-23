import pandas as pd 
import requests
import json
import os
from dotenv import load_dotenv
import urllib3

# Deshabilitar las advertencias de SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Cargar las variables del archivo .env
load_dotenv()

# Configuración de la API de GLPI
GLPI_URL = os.getenv("GLPI_URL")
USER_TOKEN = os.getenv("USER_TOKEN")
APP_TOKEN = os.getenv("APP_TOKEN")

# Validar que las variables estén configuradas
if not GLPI_URL or not USER_TOKEN or not APP_TOKEN:
    raise ValueError("Las variables GLPI_URL, USER_TOKEN o APP_TOKEN no están definidas correctamente.")

# Función para obtener el token de sesión
def obtener_token_sesion():
    headers = {
        "Authorization": f"user_token {USER_TOKEN}",
        "App-Token": APP_TOKEN,
    }
    response = requests.get(f"{GLPI_URL}/initSession", headers=headers, verify=False)
    if response.status_code == 200:
        return response.json().get("session_token")
    else:
        print(f"Error al iniciar sesión: {response.status_code}")
        print(f"Detalles del error: {response.json()}")
        return None

# Probar la conexión
session_token = obtener_token_sesion()
if session_token:
    print(f"Token de sesión obtenido: {session_token}")
else:
    print("No se pudo obtener el token de sesión.")

# Función para registrar un asset en GLPI
def registrar_asset(session_token, asset_data):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }
    # Determinar el tipo de asset según una lógica básica
    asset_type = "Network Equipment" if "network" in asset_data["name"].lower() else "Computer"
    endpoint = {
        "Computer": "/Computer",
        "Network Equipment": "/NetworkEquipment",
    }.get(asset_type, "/Computer")

    # Convertir el objeto en un array, asegurándote de que sea un objeto válido según la API
    asset_data_array = {"input": [asset_data]}  # Estructura correcta para la API

    # Enviar los datos a GLPI
    response = requests.post(f"{GLPI_URL}{endpoint}", headers=headers, data=json.dumps(asset_data_array), verify=False)
    if response.status_code == 201:
        print(f"Asset registrado exitosamente: {asset_data['name']}")
    else:
        print(f"Error al registrar asset: {response.status_code}")
        try:
            print(response.json())
        except json.JSONDecodeError:
            print(response.text)

# Función principal para procesar el archivo Excel
def procesar_archivo_excel(ruta_archivo):
    # Leer el archivo Excel desde la celda A3
    df = pd.read_excel(ruta_archivo, skiprows=2)  # Saltar las primeras 2 filas
    df.columns = df.columns.str.strip()  # Limpiar nombres de columnas

    # Verificar si las columnas necesarias están presentes
    columnas_necesarias = ["Componente", "Código", "Marca", "Location elec. Rack", "Comentario"]
    for columna in columnas_necesarias:
        if columna not in df.columns:
            print(f"Error: La columna '{columna}' no existe en el archivo Excel.")
            return

    # Obtener el token de sesión de GLPI
    session_token = obtener_token_sesion()
    if not session_token:
        print("No se pudo obtener el token de sesión.")
        return

    # Procesar solo la primera fila del archivo
    first_row = df.iloc[0]  # Obtener la primera fila
    asset_data = {
        "name": first_row["Componente"],
        "serial": first_row["Código"],
        "manufacturers": first_row["Marca"],
        "locations_id": first_row["Location elec. Rack"], #45
        "comments": first_row.get("Comentario", ""),
    }
    print(f"Procesando: {asset_data}")
    registrar_asset(session_token, asset_data)

# Ruta del archivo Excel
ruta_archivo = "C:/Users/sebas/Desktop/GLPI-Asset-Automator/Inventario Rittal_SCO y OSF_16 Nov 2021.xlsx"

# Ejecutar el proceso
procesar_archivo_excel(ruta_archivo)
