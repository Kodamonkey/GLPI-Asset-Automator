import pandas as pd
import cv2
from pyzbar.pyzbar import decode
import requests
import json
import os
from dotenv import load_dotenv
from openpyxl import load_workbook
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

# Ruta del archivo Excel
ruta_excel = "inventario_componentes.xlsx"

# Crear archivo Excel si no existe
if not os.path.exists(ruta_excel):
    columnas_necesarias = ["Código", "Componente", "Marca", "Ubicación", "Comentarios"]
    df = pd.DataFrame(columns=columnas_necesarias)
    df.to_excel(ruta_excel, index=False)

# Función para obtener el token de sesión de GLPI
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

# Función para registrar un asset en GLPI
def registrar_asset_glpi(session_token, asset_data):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }
    asset_data_array = {"input": [asset_data]}  # Estructura correcta para la API
    response = requests.post(f"{GLPI_URL}/Computer", headers=headers, data=json.dumps(asset_data_array), verify=False)
    if response.status_code == 201:
        print(f"Asset registrado exitosamente en GLPI: {asset_data['name']}")
    else:
        print(f"Error al registrar asset en GLPI: {response.status_code}")
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
def agregar_a_excel(datos):
    try:
        workbook = load_workbook(ruta_excel)
        sheet = workbook.active
        nueva_fila = [datos["Código"], datos["Componente"], datos["Marca"], datos["Ubicación"], datos["Comentarios"]]
        sheet.append(nueva_fila)
        workbook.save(ruta_excel)
        print("Datos registrados exitosamente en el Excel.")
    except Exception as e:
        print(f"Error al guardar los datos: {e}")

# Función principal
def main():
    session_token = obtener_token_sesion()
    if not session_token:
        print("No se pudo obtener el token de sesión de GLPI.")
        return

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

    # Simular búsqueda de datos (puedes integrar una lógica real aquí)
    datos = {
        "Código": codigo,
        "Componente": "Componente Desconocido",
        "Marca": "Marca Desconocida",
        "Ubicación": "Ubicación Desconocida",
        "Comentarios": "Sin comentarios"
    }

    print(f"Procesando: {datos}")
    agregar_a_excel(datos)

    # Registrar en GLPI
    asset_data = {
        "name": datos["Componente"],
        "serial": datos["Código"],
        "manufacturer": datos["Marca"],
        "locations_id": datos["Ubicación"],
        "comments": datos["Comentarios"],
    }
    registrar_asset_glpi(session_token, asset_data)

# Ejecutar la función principal
if __name__ == "__main__":
    main()
