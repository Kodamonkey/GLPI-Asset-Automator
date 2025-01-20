import pandas as pd
import cv2  # Para la captura de QR
from pyzbar.pyzbar import decode  # Decodificar QR
import os
from openpyxl import load_workbook
import requests
import json
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

# Ruta del archivo Excel
ruta_excel = "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Excel-tests.xlsx"

# Crear archivo Excel si no existe
if not os.path.exists(ruta_excel):
    columnas_necesarias = ["Asset Type", "Name", "Location", "Manufacturer", "Model", "Serial Number", 
                           "Inventory Number", "Comments", "Technician in Charge", "Group in Charge", "Status"]
    df = pd.DataFrame(columns=columnas_necesarias)
    df.to_excel(ruta_excel, index=False)

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
        return None

def obtener_location_id(session_token, location_name):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }
    
    params = {'searchText': location_name, 'range': '0-999'}
    response = requests.get(f"{GLPI_URL}/Location", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        locations = response.json()
        for location in locations:
            if location.get("name", "").strip().lower() == location_name.strip().lower():
                return location["id"]
    return None

def obtener_manufacturer_id(session_token, manufacturer_name):
    print(manufacturer_name)
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }
    
    params = {'searchText': manufacturer_name, 'range': '0-999'}
    response = requests.get(f"{GLPI_URL}/Manufacturer", headers=headers, params=params, verify=False)
    

    if response.status_code == 200:
        manufacturers = response.json()
        cleaned_manufacturer_name = manufacturer_name.strip().lower().replace("\n", "").replace("\r", "")
        
        #print("Fabricantes encontrados en GLPI:")
        #for manufacturer in manufacturers:
        #    print(f"- {manufacturer['name']} (ID: {manufacturer['id']})")

        for manufacturer in manufacturers:
            glpi_name = manufacturer.get("name", "").strip().lower().replace("\n", "").replace("\r", "")
            if glpi_name == cleaned_manufacturer_name:
                manufacturer_id = manufacturer["id"]
                print(f"ID del fabricante encontrado '{manufacturer_name}': {manufacturer_id}")
                return manufacturer_id
        
        print(f"No se encontró una coincidencia exacta para el fabricante '{manufacturer_name}'.")
        return None
    else:
        print(f"Error al obtener el fabricante: {response.status_code}")
        try:
            print(response.json())
        except json.JSONDecodeError:
            print(response.text)
        return None


def registrar_asset(session_token, asset_data, asset_type):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    endpoint = {
        "Computer": "/Computer",
        "Network Equipment": "/NetworkEquipment",
        "Consumables": "/ConsumableItem",
    }.get(asset_type, "/Computer")
    
    asset_data_array = {"input": [asset_data]}
    response = requests.post(f"{GLPI_URL}{endpoint}", headers=headers, data=json.dumps(asset_data_array), verify=False)
    
    if response.status_code == 201:
        print(f"Asset registrado exitosamente: {asset_data['name']}")
    else:
        print(f"Error al registrar asset: {response.status_code}")
        try:
            print(response.json())
        except json.JSONDecodeError:
            print(response.text)

# Función para escanear QR usando la cámara
def escanear_qr():
    cap = cv2.VideoCapture(0)
    print("Apunta la cámara al código QR. Presiona 'q' para salir.")

    while True:
        ret, frame = cap.read()
        if not ret:
            print("No se pudo acceder a la cámara.")
            break

        gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        qr_codes = decode(gray_frame)

        for qr in qr_codes:
            qr_data = qr.data.decode('utf-8')
            print(f"Código QR escaneado: \n{qr_data}")
            cap.release()
            cv2.destroyAllWindows()
            return qr_data

        cv2.imshow("Escaneando QR", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    return None

# Función para convertir la cadena del QR a un diccionario
def parse_qr_data(qr_string):
    asset_data = {}
    for line in qr_string.split("\n"):
        key, value = line.split(": ", 1)
        asset_data[key.strip()] = value.strip().replace('"', '')
    return asset_data

# Función para agregar datos al Excel
def agregar_a_excel(asset_data):
    try:
        df = pd.read_excel(ruta_excel)

        # Convertir el asset_data a un DataFrame de una fila
        nuevo_registro = pd.DataFrame([asset_data])

        # Agregar la nueva fila al DataFrame existente
        df = pd.concat([df, nuevo_registro], ignore_index=True)

        # Guardar el DataFrame actualizado en el Excel
        df.to_excel(ruta_excel, index=False)
        print("Datos registrados exitosamente en el Excel.")
    except Exception as e:
        print(f"Error al guardar los datos: {e}")

def procesar_archivo_excel(ruta_archivo):
    df = pd.read_excel(ruta_archivo, skiprows=0)
    df.columns = df.columns.str.strip()

    columnas_necesarias = [
        "Asset Type", "Name", "Location", "Manufacturer", "Model", "Serial Number", 
        "Inventory Number", "Comments", "Technician in Charge", 
        "Group in Charge", "Status", "Specific Fields (Dynamic Column)"
    ]

    for columna in columnas_necesarias:
        if columna not in df.columns:
            print(f"Error: La columna '{columna}' no existe en el archivo Excel.")
            return

    df = df.fillna("").astype(str)

    session_token = obtener_token_sesion()
    if not session_token:
        print("No se pudo obtener el token de sesión.")
        return

    # Mapeo para normalizar los tipos de assets
    asset_type_mapping = {
        "computer": "Computer",
        "network equipment": "Network Equipment",
        "consumables": "Consumables",
    }

    for index, row in df.iterrows():
        asset_type = row["Asset Type"].strip().lower()
        asset_type = asset_type_mapping.get(asset_type, None)

        if not asset_type:
            print(f"Tipo de asset desconocido: '{row['Asset Type']}' (Fila {index + 1}). Se omite esta fila.")
            continue  # Saltar la fila si el tipo de asset no es válido

        # Obtener location_id
        location_id = obtener_location_id(session_token, row["Location"].strip())
        if location_id is None:
            print(f"No se pudo encontrar la ubicación: {row['Location']} (Fila {index + 1})")
            continue  # Saltar la fila si no se encuentra la ubicación

        # Obtener manufacturer_id
        manufacturer_id = obtener_manufacturer_id(session_token, row["Manufacturer"].strip())
        print(f"ID del fabricante encontrado '{row['Manufacturer']}': {manufacturer_id}")
        if manufacturer_id is None:
            print(f"No se pudo encontrar el fabricante: {row['Manufacturer']} (Fila {index + 1})")
            continue  # Saltar la fila si no se encuentra el fabricante

        asset_data = {
            "name": row["Name"].strip(),
            "locations_id": location_id, 
            "manufacturers_id": manufacturer_id,
            "serial": row["Serial Number"].strip(),
            "otherserial": row["Inventory Number"].strip(),
            "comments": row["Comments"].strip(),
        }

        print(f"Procesando fila {index + 1}: {asset_data} como {asset_type}")
        registrar_asset(session_token, asset_data, asset_type)

# Función principal actualizada
def main():
    print("¿Deseas escanear un QR o ingresar un número de serie manualmente?")
    opcion = input("Escribe 'QR' para escanear o 'manual' para ingresar: ").strip().lower()
    
    if opcion == "qr":
        codigo = escanear_qr()
    elif opcion == "manual":
        codigo = input("Ingresa los datos manualmente: ").strip()
    else:
        print("Opción no válida.")
        return

    if not codigo:
        print("No se detectó ningún código.")
        return

    # Convertir los datos del QR en un diccionario
    asset_data = parse_qr_data(codigo)

    if asset_data:
        # Agregar los datos al archivo Excel
        agregar_a_excel(asset_data)

        # Preguntar si se desea registrar en GLPI
        registrar_glpi = input("¿Deseas registrar este activo en GLPI? (sí/no): ").strip().lower()
        if registrar_glpi == "sí" or registrar_glpi == "si":
            procesar_archivo_excel(ruta_excel)
        else:
            print("El activo no fue registrado en GLPI.")

    else:
        print("No se encontraron datos asociados a ese código.")

if __name__ == "__main__":
    main()
