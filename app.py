import pandas as pd
import cv2  # Para la captura de QR
from pyzbar.pyzbar import decode  # Decodificar QR
import os
import requests
import json
from dotenv import load_dotenv
import urllib3
import re
import numpy as np


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
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }
    
    params = {'searchText': manufacturer_name, 'range': '0-999'}
    response = requests.get(f"{GLPI_URL}/Manufacturer", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        manufacturers = response.json()
        for manufacturer in manufacturers:
            if manufacturer.get("name", "").strip().lower() == manufacturer_name.strip().lower():
                return manufacturer["id"]
    return None

def limpiar_asset_data(asset_data):
    cleaned_data = {}
    for key, value in asset_data.items():
        # Reemplazar NaN con una cadena vacía o un valor por defecto
        if isinstance(value, (float, np.float64)) and np.isnan(value):
            cleaned_data[key] = ""
        elif value is None:
            cleaned_data[key] = ""
        else:
            cleaned_data[key] = value
    return cleaned_data

def verificar_existencia_asset(session_token, serial_number):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    params = {"searchText": serial_number, "range": "0-999"}
    response = requests.get(f"{GLPI_URL}/search/Computer", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        assets = response.json().get("data", [])
        for asset in assets:
            if asset.get('5', '').strip() == serial_number:
                print(f"El activo con número de serie '{serial_number}' ya existe en GLPI con ID: {asset.get('1')}")
                return True
    else:
        print(f"Error al verificar existencia del activo: {response.status_code}")
        try:
            print(response.json())
        except json.JSONDecodeError:
            print(response.text)
    
    return False

def registrar_asset(session_token, asset_data, asset_type):
    if verificar_existencia_asset(session_token, asset_data["serial"]):
        print(f"El activo con número de serie {asset_data['serial']} ya existe en GLPI. No se realizará el registro.")
        return

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

    # Limpiar datos antes de enviarlos
    asset_data_clean = limpiar_asset_data(asset_data)

    # Crear la estructura correcta para la API de GLPI
    asset_data_array = {"input": [asset_data_clean]}
    
    response = requests.post(f"{GLPI_URL}{endpoint}", headers=headers, data=json.dumps(asset_data_array), verify=False)
    
    if response.status_code == 201:
        print(f"Asset registrado exitosamente: {asset_data_clean['name']}")
    else:
        print(f"Error al registrar asset: {response.status_code}")
        try:
            print(response.json())
        except json.JSONDecodeError:
            print(response.text)

def registrar_ultima_fila():
    df = pd.read_excel(ruta_excel)
    if df.empty:
        print("El archivo Excel está vacío.")
        return

    last_row = df.iloc[-1].to_dict()
    print(f"Última fila encontrada: {last_row}")

    session_token = obtener_token_sesion()
    if not session_token:
        print("No se pudo obtener el token de sesión.")
        return

    # Verificar si 'name' existe en la fila
    if "Name" not in last_row or "Asset Type" not in last_row:
        print("La última fila no contiene las columnas esperadas.")
        return

    location_id = obtener_location_id(session_token, last_row["Location"])
    if not location_id:
        print(f"No se pudo encontrar la ubicación: {last_row['Location']}")
        return
    manufacturer_id = obtener_manufacturer_id(session_token, last_row["Manufacturer"])
    if not manufacturer_id:
        print(f"No se pudo encontrar el fabricante: {last_row['Manufacturer']}")
        return

    if location_id is None or manufacturer_id is None:
        print(f"No se pudo encontrar la ubicación o el fabricante para el activo '{last_row['Name']}'")
        return

    # Preparar los datos para el registro en GLPI
    asset_data = {
        "name": last_row["Name"].strip(),
        "locations_id": location_id,
        "manufacturers_id": manufacturer_id,
        "serial": last_row["Serial Number"].strip(),
        #"otherserial": last_row["Inventory Number"].strip(),
        "comments": last_row["Comments"].strip(),
    }

    print(f"Registrando asset: {asset_data}")

    registrar_asset(session_token, asset_data, last_row["Asset Type"])

def registrar_por_nombre():
    df = pd.read_excel(ruta_excel)
    if df.empty:
        print("El archivo Excel está vacío.")
        return

    nombre = input("Ingrese el nombre del activo a registrar: ").strip()
    filtro = df[df["Name"].str.lower() == nombre.lower()]

    if filtro.empty:
        print(f"No se encontró el activo con el nombre '{nombre}' en el archivo Excel.")
        return

    row = filtro.iloc[0].to_dict()
    session_token = obtener_token_sesion()

    location_id = obtener_location_id(session_token, row["Location"])
    manufacturer_id = obtener_manufacturer_id(session_token, row["Manufacturer"])

    if location_id is None or manufacturer_id is None:
        print(f"No se pudo encontrar la ubicación o el fabricante para el activo '{row['Name']}'")
        return

    # Preparar los datos para el registro en GLPI
    asset_data = {
        "name": row["Name"].strip(),
        "locations_id": location_id,
        "manufacturers_id": manufacturer_id,
        "serial": row["Serial Number"].strip(),
        #"otherserial": row["Inventory Number"].strip(),
        "comments": row["Comments"].strip(),
    }

    registrar_asset(session_token, asset_data, row["Asset Type"])

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
            #"otherserial": row["Inventory Number"].strip(),
            "comments": row["Comments"].strip(),
        }

        print(f"Procesando fila {index + 1}: {asset_data} como {asset_type}")
        registrar_asset(session_token, asset_data, asset_type)

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

def escanear_qr_con_celular():
    ip_cam_url = "https://10.200.253.178:8080/video"  # Cambiar por la URL de la cámara IP

    while True:
        cap = cv2.VideoCapture(ip_cam_url)
        if not cap.isOpened():
            print("No se pudo acceder a la cámara del celular. Reintentando en 5 segundos...")
            cap.release()
            cv2.destroyAllWindows()
            cv2.waitKey(5000)  # Esperar 5 segundos antes de reintentar
            continue

        print("Usando la cámara del celular. Presiona 'q' para salir.")
        
        while True:
            ret, frame = cap.read()
            if not ret:
                print("Error al obtener el cuadro de la cámara. Reintentando conexión...")
                break  # Sale del bucle interno para reintentar la conexión
            
            gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            qr_codes = decode(gray_frame)

            for qr in qr_codes:
                qr_data = qr.data.decode('utf-8')
                print(f"Código QR escaneado: \n{qr_data}")
                cap.release()
                cv2.destroyAllWindows()
                return qr_data

            cv2.imshow("Escaneando QR con celular", frame)

            if cv2.waitKey(1) & 0xFF == ord('q'):
                cap.release()
                cv2.destroyAllWindows()
                return None

        cap.release()
        cv2.destroyAllWindows()

def parse_qr_data(qr_string):
    asset_data = {}
    for line in qr_string.split("\n"):
        key, value = line.split(": ", 1)
        asset_data[key.strip()] = value.strip().replace('"', '')
    return asset_data

def verificar_existencia_en_excel(serial_number):
    df = pd.read_excel(ruta_excel)
    if serial_number in df["Serial Number"].values:
        print(f"El activo con número de serie '{serial_number}' ya existe en el Excel.")
        return True
    return False

def agregar_a_excel(asset_data):
    try:
        df = pd.read_excel(ruta_excel)

        if verificar_existencia_en_excel(asset_data["Serial Number"]):
            print(f"El activo con serial '{asset_data['Serial Number']}' ya está registrado en el Excel. No se agregará.")
            return

        nuevo_registro = pd.DataFrame([asset_data]) # Convertir el asset_data a un DataFrame de una fila
        df = pd.concat([df, nuevo_registro], ignore_index=True) # Agregar la nueva fila al DataFrame existente
        df.to_excel(ruta_excel, index=False) # Guardar el DataFrame actualizado en el Excel
        print("Datos registrados exitosamente en el Excel.")

        # Guardar la plantilla en un archivo .txt
        nombre_archivo_txt = f"C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Templates/{asset_data['Name']}.txt"
        guardar_plantilla_txt(asset_data, nombre_archivo_txt)

        # Preguntar si se desea registrar en GLPI
        registrar_glpi = input("¿Deseas registrar este activo en GLPI? (sí/no): ").strip().lower()
        if registrar_glpi == "sí" or registrar_glpi == "si":
            registrar_ultima_fila()
        else:
            print("El activo no fue registrado en GLPI.")
    except Exception as e:
        print(f"Error al guardar los datos: {e}")

def procesar_qr_dell(qr_data):
    # Plantilla de la laptop Dell Latitude
    plantilla_dell = {
        "Asset Type": "Computer",
        "Status": "Stocked",  # Solicitar por pantalla
        "User": None,    # Solicitar por pantalla
        "Name": None,    # Generado a partir del nombre del usuario
        "Computer Types": "Laptop",
        "Location": None,  # Solicitar por pantalla
        "Manufacturer": "Dell inc.",
        "Model": "Latitude",
        "Serial Number": qr_data.strip(),  # QR escaneado de la laptop
        "Comments": "Check",
    }

    # Solicitar datos adicionales al usuario
    #plantilla_dell["Status"] = input("Ingrese el estado del activo (Activo/Inactivo): ").strip()
    #plantilla_dell["User"] = input("Ingrese el nombre del usuario: ").strip()
    plantilla_dell["Location"] = input("Ingrese la ubicación del activo: ").strip()
    
    # Generar el nombre del activo a partir del usuario
    plantilla_dell["Name"] = f"{plantilla_dell['User']}-Latitude"

    return plantilla_dell

def procesar_qr_mac(qr_data):
    # Plantilla para laptops MacBook
    plantilla_mac = {
        "Asset Type": "Computer",
        "Status": "Stocked",  # Solicitar por pantalla
        "User": None,    # Solicitar por pantalla
        "Name": None,    # Generado a partir del nombre del usuario
        "Computer Types": "Laptop",
        "Location": None,  # Solicitar por pantalla
        "Manufacturer": "Apple Inc",
        "Model": "MacBook Pro",
        "Serial Number": qr_data.strip(),  # QR escaneado de la laptop
        "Comments": "Check",
    }

    # Solicitar datos adicionales al usuario
    #plantilla_mac["Status"] = input("Ingrese el estado del activo (Activo/Inactivo): ").strip()
    #plantilla_mac["User"] = input("Ingrese el nombre del usuario: ").strip()
    plantilla_mac["Location"] = input("Ingrese la ubicación del activo: ").strip()
    
    # Generar el nombre del activo a partir del usuario
    plantilla_mac["Name"] = f"{plantilla_mac['User']}-MacBookPro"

    return plantilla_mac

def guardar_plantilla_txt(asset_data, nombre_archivo):
    with open(nombre_archivo, 'w') as file:
        for key, value in asset_data.items():
            file.write(f"{key}: {value}\n")
    print(f"Plantilla guardada en {nombre_archivo}")

def agregar_a_excel_dell(asset_data):
    try:
        df = pd.read_excel(ruta_excel)
        nuevo_registro = pd.DataFrame([asset_data])  # Convertir la plantilla a un DataFrame de una fila
        df = pd.concat([df, nuevo_registro], ignore_index=True)
        df.to_excel(ruta_excel, index=False)
        print("Datos registrados exitosamente en el Excel.")

        # Guardar la plantilla en un archivo .txt
        nombre_archivo_txt = f"C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Templates/{asset_data['Name']}.txt"
        guardar_plantilla_txt(asset_data, nombre_archivo_txt)

    except Exception as e:
        print(f"Error al guardar los datos: {e}")

def agregar_a_excel_mac(asset_data):
    try:
        if verificar_existencia_en_excel(asset_data["Serial Number"]):
            print(f"El activo con serial '{asset_data['Serial Number']}' ya está registrado en el Excel. No se agregará.")
            return
        df = pd.read_excel(ruta_excel)
        nuevo_registro = pd.DataFrame([asset_data])  # Convertir la plantilla a un DataFrame de una fila
        df = pd.concat([df, nuevo_registro], ignore_index=True)
        df.to_excel(ruta_excel, index=False)
        print("Datos registrados exitosamente en el Excel.")

        # Guardar la plantilla en un archivo .txt
        nombre_archivo_txt = f"C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Templates/{asset_data['Name']}.txt"
        guardar_plantilla_txt(asset_data, nombre_archivo_txt)

    except Exception as e:
        print(f"Error al guardar los datos: {e}")
        
def extraer_service_tag(qr_data):
    match = re.search(r"Service tag:\s*([A-Za-z0-9]+)", qr_data, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None

def extraer_serial_mac(qr_data):
    match = re.search(r"Serial Number:\s*([A-Za-z0-9]+)", qr_data, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    
    # Patrón alternativo para seriales de Mac que comienzan con 'C02' y tienen entre 10 y 12 caracteres
    match_alt = re.search(r"\b(C02[A-Za-z0-9]{8,10})\b", qr_data)
    if match_alt:
        return match_alt.group(1).strip()

    return None

def manejar_qr_dell():
    qr_data = escanear_qr_con_celular()
    if qr_data:
        # Mejorar la detección para considerar patrones adicionales
        if any(keyword in qr_data.lower() for keyword in ["dell", "service tag", "made in vietnam", "Service tag", "Dell", "S/N", "(S/N)", "SN"]) or qr_data.startswith("CS") or len(qr_data) == 7:
            print("Laptop Dell detectada. Procesando datos...")
            
            if len(qr_data) > 7:  # Verificar si el QR escaneado contiene más información de la cuenta
                service_tag = extraer_service_tag(qr_data)
            else:
                service_tag = qr_data  # Si solo tiene el serial en QR directamente
            
            if service_tag:
                print(f"Service Tag detectado: {service_tag}")
                confirmacion = input("¿Es correcto este Service Tag, desea continuar? (sí/no): ").strip().lower()
                if confirmacion not in ["sí", "si", "Si", "Sí"]:
                    print("Operación cancelada por el usuario.")
                    return
                else:
                    if verificar_existencia_en_excel(service_tag):
                        print(f"El activo con serial '{service_tag}' ya está registrado en el Excel. No se agregará.")
                        return
                    asset_data = procesar_qr_dell(service_tag)
                    agregar_a_excel_dell(asset_data)
            else:
                print("No se detectó un Service Tag válido en el QR escaneado.")
        else:
            print("Código QR no corresponde a un equipo Dell.")
    else:
        print("No se detectó ningún código QR.")

def manejar_qr_mac():
    qr_data = escanear_qr_con_celular()
    if qr_data:
        if "MacBook" in qr_data or "Serial Number:" in qr_data or qr_data.startswith("C02") or 10<= len(qr_data) <= 12:  # Detectar si es Mac por patrones comunes
            print("Laptop MacBook detectada. Procesando datos...")

            if len(qr_data) > 12:
                serial_number = extraer_serial_mac(qr_data)
            else:
                serial_number = qr_data

            if serial_number:
                print(f"Serial Number detectado: {serial_number}")
                confirmacion = input("¿Es correcto este Service Tag, desea continuar? (sí/no): ").strip().lower()
                if confirmacion not in ["sí", "si", "Si", "Sí"]:
                    print("Operación cancelada por el usuario.")
                    return
                else:
                    if verificar_existencia_en_excel(serial_number):
                        print(f"El activo con serial '{serial_number}' ya está registrado en el Excel. No se agregará.")
                        return
                    asset_data = procesar_qr_mac(serial_number)
                    agregar_a_excel_mac(asset_data)
        else:
            print("Código QR no corresponde a un equipo Mac.")
    else:
        print("No se detectó ningún código QR.")

def manejar_qr_laptop():
    while True:
        qr_data = escanear_qr_con_celular()

        if qr_data:
            qr_data_lower = qr_data.lower()

            # Detectar si es una laptop Dell
            patrones_dell = ["dell", "service tag", "made in vietnam", "Service tag", "Dell", "S/N", "(S/N)", "SN"]
            if any(keyword in qr_data_lower for keyword in patrones_dell) or qr_data.startswith("CS") or len(qr_data) == 7:
                print("Laptop Dell detectada. Procesando datos...")

                if len(qr_data) > 7:
                    serial_number = extraer_service_tag(qr_data)
                else:
                    serial_number = qr_data  # Si el QR contiene solo el serial

                if serial_number:
                    print(f"Service Tag detectado: {serial_number}")
                    confirmacion = input("¿Es correcto este Service Tag, desea continuar? (sí/no): ").strip().lower()
                    if confirmacion in ["sí", "si", "Si", "Sí"]:
                        if verificar_existencia_en_excel(serial_number):
                            print(f"El activo con serial '{serial_number}' ya está registrado en el Excel. No se agregará.")
                            return
                        asset_data = procesar_qr_dell(serial_number)
                        agregar_a_excel_dell(asset_data)
                        break
                    else:
                        print("Reintentando escaneo...")
                        continue
                else:
                    print("No se detectó un Service Tag válido. Reintentando...")
                    continue

            # Detectar si es una laptop MacBook
            if any(keyword in qr_data_lower for keyword in ["macbook", "serial number"]) or qr_data.startswith("C02") or 10 <= len(qr_data) <= 12:
                print("Laptop MacBook detectada. Procesando datos...")

                if len(qr_data) > 12:
                    serial_number = extraer_serial_mac(qr_data)
                else:
                    serial_number = qr_data

                if serial_number:
                    print(f"Serial Number detectado: {serial_number}")
                    confirmacion = input("¿Es correcto este Serial Number, desea continuar? (sí/no): ").strip().lower()
                    if confirmacion in ["sí", "si", "Si", "Sí"]:
                        if verificar_existencia_en_excel(serial_number):
                            print(f"El activo con serial '{serial_number}' ya está registrado en el Excel. No se agregará.")
                            return
                        asset_data = procesar_qr_mac(serial_number)
                        agregar_a_excel_mac(asset_data)
                        break
                    else:
                        print("Reintentando escaneo...")
                        continue
                else:
                    print("No se detectó un Serial Number válido. Reintentando...")
                    continue

            print("Código QR no corresponde a un equipo Dell ni Mac. Reintentando...")
        else:
            print("No se detectó ningún código QR. Reintentando...")

def entregar_laptop():
    print("\n--- Entregar Laptop a Usuario ---")
    metodo = input("¿Desea escanear el QR o ingresar el Service Tag manualmente? (escanear/manual): ").strip().lower()

    if metodo == "escanear":
        qr_data = escanear_qr_con_celular()
        if any(keyword in qr_data.lower() for keyword in ["dell", "service tag", "made in vietnam", "s/n", "(s/n)", "sn"]) or qr_data.startswith("CS") or len(qr_data) == 7:
            print("Laptop Dell detectada. Procesando datos...")
            serial_number = extraer_service_tag(qr_data) if len(qr_data) > 7 else qr_data
        elif any(keyword in qr_data.lower() for keyword in ["macbook", "serial number"]) or qr_data.startswith("C02") or 10 <= len(qr_data) <= 12:
            print("Laptop MacBook detectada. Procesando datos...")
            serial_number = extraer_serial_mac(qr_data) if len(qr_data) > 12 else qr_data
        else:
            print("Código QR no corresponde a un equipo Dell ni Mac.")
            return
    elif metodo == "manual":
        serial_number = input("Ingrese el Service Tag del laptop: ").strip()
    else:
        print("Método no válido. Intente nuevamente.")
        return

    df = pd.read_excel(ruta_excel)
    if df.empty:
        print("El archivo Excel está vacío.")
        return

    filtro = df[df["Serial Number"].str.lower() == serial_number.lower()]

    if filtro.empty:
        print(f"No se encontró un laptop con el Service Tag '{serial_number}' en el archivo Excel.")
        return

    nuevo_usuario = input("Ingrese el nombre del usuario que recibirá el laptop: ").strip()

    # Manejar valores NaN antes de actualizar el DataFrame
    df["User"] = df["User"].fillna("")
    df["Name"] = df["Name"].fillna("Unknown")

    if "Dell" in filtro["Manufacturer"].values[0]:
        name_laptop = f"{nuevo_usuario}-Latitude"
    elif "Apple" in filtro["Manufacturer"].values[0]:
        name_laptop = f"{nuevo_usuario}-MacBookPro"
    else:
        print("No se pudo determinar el fabricante del laptop.")
        return

    df.loc[df["Serial Number"].str.lower() == serial_number.lower(), "User"] = nuevo_usuario
    df.loc[df["Serial Number"].str.lower() == serial_number.lower(), "Name"] = name_laptop

    # Guardar los cambios en el Excel
    df.to_excel(ruta_excel, index=False)
    print(f"Laptop con Service Tag '{serial_number}' asignado a '{nuevo_usuario}' en el Excel.")

    # Actualizar en GLPI
    session_token = obtener_token_sesion()
    if not session_token:
       print("No se pudo obtener el token de sesión.")
       return

    #asset_id = obtener_detalles_asset(session_token, serial_number)
    #if not asset_id:
    #    print("No se pudo actualizar el activo en GLPI porque no se encontró.")
    #    return

    #asset_data = filtro.iloc[0].to_dict()
    #asset_data["User"] = nuevo_usuario
    #asset_data["Name"] = name_laptop  # Asegurar que el nombre esté presente

    # Limpiar NaN antes de enviar a GLPI
    #asset_data_clean = {key: value if pd.notna(value) else "" for key, value in asset_data.items()}

    #if "Name" not in asset_data_clean or not asset_data_clean["Name"]:
    #    print("Error: No se pudo determinar el nombre del laptop para el registro.")
    #    return
    #obtener_detalles_asset(session_token, asset_id)
    #actualizar_asset(session_token, asset_id, asset_data_clean)

# Agregar la opción en el menú principal
def main():
    while True:
        print("\n--- Menú de opciones ---")
        print("0. Escanear y registrar cualquier laptop (Dell/Mac), !Me siento con suerte!")
        print("1. Escanear QR y registrar en Excel (Template Default)")
        print("2. Escanear QR y registrar laptops Dell")
        print("3. Escanear QR y registrar laptops Mac")
        print("4. Registrar la última fila del Excel en GLPI")
        print("5. Registrar un activo por nombre")
        print("6. Registrar todos los activos de Excel en GLPI")
        print("7. Entregar laptop a un usuario")
        print("8. Salir")
        
        opcion = input("Seleccione una opción: ").strip()
        
        if opcion == "0":
            manejar_qr_laptop()
        elif opcion == "1":
            codigo = escanear_qr_con_celular()
            if codigo:
                asset_data = parse_qr_data(codigo)
                agregar_a_excel(asset_data)
        elif opcion == "2":
            manejar_qr_dell()
        elif opcion == "3":
            manejar_qr_mac()
        elif opcion == "4":
            registrar_ultima_fila()
        elif opcion == "5":
            registrar_por_nombre()
        elif opcion == "6":
            procesar_archivo_excel(ruta_excel)
        elif opcion == "7":
            entregar_laptop()
        elif opcion == "8":
            print("Saliendo del programa...")
            break
        else:
            print("Opción no válida. Intente nuevamente.")

if __name__ == "__main__":
    main()
