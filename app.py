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
ruta_excel = "C:/Users/sebas/Desktop/GLPI-Asset-Automator/Inventario Rittal_SCO y OSF_16 Nov 2021.xlsx"

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
    
    params = {'searchText': location_name}
    response = requests.get(f"{GLPI_URL}/Location", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        locations = response.json()
        for location in locations:
            if location.get("name", "").strip().lower() == location_name.strip().lower():
                return location["id"]
    return None

def obtener_manufacturer_id(session_token, manufacturer_name):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }
    
    params = {'searchText': manufacturer_name}
    response = requests.get(f"{GLPI_URL}/Manufacturer", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        manufacturers = response.json()
        for manufacturer in manufacturers:
            if manufacturer.get("name", "").strip().lower() == manufacturer_name.strip().lower():
                return manufacturer["id"]
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
            print(f"Código QR escaneado: {qr_data}")
            cap.release()
            cv2.destroyAllWindows()
            return qr_data

        cv2.imshow("Escaneando QR", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    return None

# Función para agregar datos al Excel
def agregar_a_excel(dato):
    try:
        workbook = load_workbook(ruta_excel)
        sheet = workbook.active
        nueva_fila = [dato["Asset Type"], dato["Name"], dato["Location"], dato["Manufacturer"], dato["Model"], 
                      dato["Serial Number"], dato["Inventory Number"], dato["Comments"], 
                      dato["Technician in Charge"], dato["Group in Charge"], dato["Status"], dato["Specific Fields (Dynamic Column)"]]
        sheet.append(nueva_fila)
        workbook.save(ruta_excel)
        print("Datos registrados exitosamente en el Excel.")
    except Exception as e:
        print(f"Error al guardar los datos: {e}")

# Función principal
def main():
    print("¿Deseas escanear un QR o ingresar un número de serie manualmente?")
    opcion = input("Escribe 'QR' para escanear o 'manual' para ingresar: ").strip().lower()
    
    if opcion == "qr":
        codigo = escanear_qr()
    elif opcion == "manual":
        codigo = input("Ingresa el número de serie: ").strip()
    else:
        print("Opción no válida.")
        return

    if not codigo:
        print("No se detectó ningún código.")
        return

    df = pd.read_excel(ruta_excel)
    asset_data = df[df["Serial Number"] == codigo].to_dict(orient="records")

    if asset_data:
        asset_data = asset_data[0]

        # Obtener session token
        session_token = obtener_token_sesion()
        if not session_token:
            print("No se pudo obtener el token de sesión.")
            return

        # Obtener IDs requeridos
        location_id = obtener_location_id(session_token, asset_data["Location"])
        manufacturer_id = obtener_manufacturer_id(session_token, asset_data["Manufacturer"])

        if not location_id or not manufacturer_id:
            print("Error al obtener ID de ubicación o fabricante.")
            return

        asset_data["locations_id"] = location_id
        asset_data["manufacturers_id"] = manufacturer_id

        registrar_asset(session_token, asset_data, asset_data["Asset Type"])
    else:
        print("No se encontraron datos asociados a ese código.")

if __name__ == "__main__":
    main()
