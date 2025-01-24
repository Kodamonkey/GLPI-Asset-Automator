import tkinter as tk
from tkinter import messagebox, simpledialog
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
                           "Inventory Number", "Comments", "Technician in Charge", "Group in Charge", "Status", "Specific Fields (Dynamic Column)"]
    df = pd.DataFrame(columns=columnas_necesarias)
    df.to_excel(ruta_excel, index=False)

# Ruta del archivo Excel
ruta_excel_consumibles = "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/consumibles.xlsx"

# Crear archivo Excel si no existe
if not os.path.exists(ruta_excel_consumibles):
    columnas_necesarias = ["Name", "Inventory/Asset Number", "Location", "Stock Target"]
    df = pd.DataFrame(columns=columnas_necesarias)
    df.to_excel(ruta_excel_consumibles, index=False)


def obtener_token_sesion():
    headers = {
        "Authorization": f"user_token {USER_TOKEN}",
        "App-Token": APP_TOKEN,
    }
    response = requests.get(f"{GLPI_URL}/initSession", headers=headers, verify=False)
    if response.status_code == 200:
        return response.json().get("session_token")
    else:
        messagebox.showerror("Error", f"Error al iniciar sesión: {response.status_code}")
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
        if isinstance(value, float) and np.isnan(value):
            cleaned_data[key] = ""
        elif value is None:
            cleaned_data[key] = ""
        else:
            cleaned_data[key] = str(value).strip()
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
                messagebox.showinfo("Información", f"El activo con número de serie '{serial_number}' ya existe en GLPI con ID: {asset.get('1')}")
                return True
    else:
        messagebox.showerror("Error", f"Error al verificar existencia del activo: {response.status_code}")
        try:
            messagebox.showerror("Error", response.json())
        except json.JSONDecodeError:
            messagebox.showerror("Error", response.text)
    
    return False

def registrar_asset(session_token, asset_data, asset_type):
    if verificar_existencia_asset(session_token, asset_data["serial"]):
        messagebox.showinfo("Información", f"El activo con número de serie {asset_data['serial']} ya existe en GLPI. No se realizará el registro.")
        return

    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    endpoint = {
        "Computer": "/Computer",
        "Monitor": "/Monitor",
        "Network Equipment": "/NetworkEquipment",
        "Consumables": "/ConsumableItem",
    }.get(asset_type, "/Computer")

    # Limpiar datos antes de enviarlos
    asset_data_clean = limpiar_asset_data(asset_data)

    # Crear la estructura correcta para la API de GLPI
    asset_data_array = {"input": [asset_data_clean]}
    
    response = requests.post(f"{GLPI_URL}{endpoint}", headers=headers, data=json.dumps(asset_data_array), verify=False)
    
    if response.status_code == 201:
        messagebox.showinfo("Éxito", f"Asset registrado exitosamente: {asset_data_clean['name']}")
    else:
        messagebox.showerror("Error", f"Error al registrar asset: {response.status_code}")
        try:
            messagebox.showerror("Error", response.json())
        except json.JSONDecodeError:
            messagebox.showerror("Error", response.text)

def registrar_ultima_fila():
    df = pd.read_excel(ruta_excel)
    if df.empty:
        messagebox.showerror("Error", "El archivo Excel está vacío.")
        return

    last_row = df.iloc[-1].to_dict()
    #messagebox.showinfo("Información", f"Última fila encontrada: {last_row}")

    # Reemplazar NaN con valores vacíos
    last_row = {key: ("" if pd.isna(value) else value) for key, value in last_row.items()}

    session_token = obtener_token_sesion()
    if not session_token:
        messagebox.showerror("Error", "No se pudo obtener el token de sesión.")
        return

    if "Name" not in last_row or "Asset Type" not in last_row:
        messagebox.showerror("Error", "La última fila no contiene las columnas esperadas.")
        return

    location_id = obtener_location_id(session_token, last_row["Location"])
    if not location_id:
        messagebox.showerror("Error", f"No se pudo encontrar la ubicación: {last_row['Location']}")
        return

    manufacturer_id = obtener_manufacturer_id(session_token, last_row["Manufacturer"])
    if not manufacturer_id:
        messagebox.showerror("Error", f"No se pudo encontrar el fabricante: {last_row['Manufacturer']}")
        return

    if location_id is None or manufacturer_id is None:
        messagebox.showerror("Error", f"No se pudo encontrar la ubicación o el fabricante para el activo '{last_row['Name']}'")
        return

    # Preparar los datos para el registro en GLPI
    asset_data = {
        "name": last_row["Name"].strip(),
        "locations_id": location_id,
        "manufacturers_id": manufacturer_id,
        "serial": last_row["Serial Number"].strip(),
        "comments": last_row["Comments"].strip() if last_row["Comments"] else "N/A",
    }

    messagebox.showinfo("Información", f"Registrando asset: {asset_data}")

    registrar_asset(session_token, asset_data, last_row["Asset Type"])

def registrar_por_nombre():
    df = pd.read_excel(ruta_excel)
    if df.empty:
        messagebox.showerror("Error", "El archivo Excel está vacío.")
        return

    nombre = simpledialog.askstring("Input", "Ingrese el nombre del activo a registrar:").strip()
    if not nombre:
        messagebox.showerror("Error", "No se ingresó un nombre válido.")
        return

    filtro = df[df["Name"].str.lower() == nombre.lower()]

    if filtro.empty:
        messagebox.showerror("Error", f"No se encontró el activo con el nombre '{nombre}' en el archivo Excel.")
        return

    row = filtro.iloc[0].to_dict()
    session_token = obtener_token_sesion()

    location_id = obtener_location_id(session_token, row["Location"])
    manufacturer_id = obtener_manufacturer_id(session_token, row["Manufacturer"])

    if location_id is None or manufacturer_id is None:
        messagebox.showerror("Error", f"No se pudo encontrar la ubicación o el fabricante para el activo '{row['Name']}'")
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
            messagebox.showerror("Error", f"La columna '{columna}' no existe en el archivo Excel.")
            return

    df = df.fillna("").astype(str)

    session_token = obtener_token_sesion()
    if not session_token:
        messagebox.showerror("Error", "No se pudo obtener el token de sesión.")
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
            messagebox.showerror("Error", f"Tipo de asset desconocido: '{row['Asset Type']}' (Fila {index + 1}). Se omite esta fila.")
            continue  # Saltar la fila si el tipo de asset no es válido

        # Obtener location_id
        location_id = obtener_location_id(session_token, row["Location"].strip())
        if location_id is None:
            messagebox.showerror("Error", f"No se pudo encontrar la ubicación: {row['Location']} (Fila {index + 1})")
            continue  # Saltar la fila si no se encuentra la ubicación

        # Obtener manufacturer_id
        manufacturer_id = obtener_manufacturer_id(session_token, row["Manufacturer"].strip())
        #messagebox.showinfo("Información", f"ID del fabricante encontrado '{row['Manufacturer']}': {manufacturer_id}")
        if manufacturer_id is None:
            messagebox.showerror("Error", f"No se pudo encontrar el fabricante: {row['Manufacturer']} (Fila {index + 1})")
            continue  # Saltar la fila si no se encuentra el fabricante

        asset_data = {
            "name": row["Name"].strip(),
            "locations_id": location_id, 
            "manufacturers_id": manufacturer_id,
            "serial": row["Serial Number"].strip(),
            #"otherserial": row["Inventory Number"].strip(),
            "comments": row["Comments"].strip(),
        }

        messagebox.showinfo("Información", f"Procesando fila {index + 1}: {asset_data} como {asset_type}")
        registrar_asset(session_token, asset_data, asset_type)

def escanear_qr():
    cap = cv2.VideoCapture(0)
    messagebox.showinfo("Información", "Apunta la cámara al código QR. Presiona 'q' para salir.")

    while True:
        ret, frame = cap.read()
        if not ret:
            messagebox.showerror("Error", "No se pudo acceder a la cámara.")
            break

        gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        qr_codes = decode(gray_frame)

        for qr in qr_codes:
            qr_data = qr.data.decode('utf-8')
            messagebox.showinfo("Información", f"Código QR escaneado: \n{qr_data}")
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
            messagebox.showerror("Error", "No se pudo acceder a la cámara del celular. Reintentalo...")
            cap.release()
            cv2.destroyAllWindows()
            cv2.waitKey(5000)  # Esperar 5 segundos antes de reintentar
            continue

        messagebox.showinfo("Información", "Usando la cámara del celular. Presiona 'q' para salir.")
        
        while True:
            ret, frame = cap.read()
            if not ret:
                messagebox.showerror("Error", "Error al obtener el cuadro de la cámara. Reintentando conexión...")
                break  # Sale del bucle interno para reintentar la conexión
            
            gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            qr_codes = decode(gray_frame)

            for qr in qr_codes:
                qr_data = qr.data.decode('utf-8')
                if es_codigo_valido(qr_data):
                    messagebox.showinfo("Información", f"Código QR escaneado: \n{qr_data}")
                    cap.release()
                    cv2.destroyAllWindows()
                    return qr_data
                else:
                    print("Código QR no válido. Reintentando...")
                    break

            cv2.imshow("Escaneando QR con celular", frame)

            if cv2.waitKey(1) & 0xFF == ord('q'):
                cap.release()
                cv2.destroyAllWindows()
                return None

        cap.release()
        cv2.destroyAllWindows()

def es_codigo_valido(qr_data):
    # lógica para determinar si un código QR es válido
    # verificar si el código QR coincide con ciertos patrones
    patrones_validos = [
        #r'\bCS[A-Z0-9]{5}\b',  # Patrón para Dell
        #r'^[A-Za-z0-9]{7}$',   # Patrón para Dell
        r'\bC02[A-Za-z0-9]{8,10}\b',  # Patrón para Mac
        r'^[A-Za-z0-9]{10}$',  # Patrón para Mac
        r'^S?[C02][A-Za-z0-9]{8}$',  # Patrón para Mac
        r'^S[A-Za-z0-9]{10}$'  # Patrón para Mac
        r'^S[A-Za-z0-9]{11}$'  # Patrón para Mac

        # Dell Service Tag: 7 caracteres alfanuméricos, excluyendo I, O y Q (para evitar confusión con números)
        r'^(?!.*[IOQ])[A-HJ-NP-Z0-9]{7}$',  # Ejemplo: 8B9X1R3
        
        # Dell Express Service Code (conversión numérica del Service Tag)
        #r'^\d{10}$',  # Ejemplo: 1234567890

        # MacBook Serial Number: Inicia con C02 o FVX seguido de 8-10 caracteres alfanuméricos
        r'^C02[A-Z0-9]{8,10}$',  # Ejemplo: C02X3Y5VFH5
        r'^FVX[A-Z0-9]{8,10}$',  # Ejemplo: FVXJ45KLD9

        # MacBook Serial Numbers con prefijo opcional 'S'
        r'^S?(C02|FVX)[A-Z0-9]{8,10}$',  # Soporta opcional 'S' delante del serial

        # MacBook Air/Pro: Formato reciente con 12 caracteres alfanuméricos
        #r'^[A-Z0-9]{12}$',  # Ejemplo: W8P6W5T5YV3C
    ]
    for patron in patrones_validos:
        if re.match(patron, qr_data):
            return True
    return False

def parse_qr_data(qr_string):
    asset_data = {}
    for line in qr_string.split("\n"):
        key, value = line.split(": ", 1)
        asset_data[key.strip()] = value.strip().replace('"', '')
    return asset_data

def verificar_existencia_en_excel(serial_number):
    df = pd.read_excel(ruta_excel)
    if serial_number in df["Serial Number"].values:
        messagebox.showinfo("Información", f"El activo con número de serie '{serial_number}' ya existe en el Excel.")
        return True
    return False

def agregar_a_excel(asset_data):
    try:
        df = pd.read_excel(ruta_excel)

        if verificar_existencia_en_excel(asset_data["Serial Number"]):
            messagebox.showinfo("Información", f"El activo con serial '{asset_data['Serial Number']}' ya está registrado en el Excel. No se agregará.")
            return

        nuevo_registro = pd.DataFrame([asset_data])  # Convertir el asset_data a un DataFrame de una fila
        df = pd.concat([df, nuevo_registro], ignore_index=True)  # Agregar la nueva fila al DataFrame existente
        df.to_excel(ruta_excel, index=False)  # Guardar el DataFrame actualizado en el Excel
        messagebox.showinfo("Éxito", "Datos registrados exitosamente en el Excel.")

        # Guardar la plantilla en un archivo .txt
        nombre_archivo_txt = f"C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Templates/{asset_data['Name']}.txt"
        guardar_plantilla_txt(asset_data, nombre_archivo_txt)

        # Preguntar si se desea registrar en GLPI
        registrar_glpi = simpledialog.askstring("Registrar en GLPI", "¿Deseas registrar este activo en GLPI? (sí/no):").strip().lower()
        if registrar_glpi == "sí" or registrar_glpi == "si":
            registrar_ultima_fila()
        else:
            messagebox.showinfo("Información", "El activo no fue registrado en GLPI.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar los datos: {e}")

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
    #plantilla_dell["Status"] = simpledialog.askstring("Input", "Ingrese el estado del activo (Activo/Inactivo):").strip()
    #plantilla_dell["User"] = simpledialog.askstring("Input", "Ingrese el nombre del usuario:").strip()
    plantilla_dell["Location"] = simpledialog.askstring("Input", "Ingrese la ubicación del activo:").strip()
    
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
    #plantilla_mac["Status"] = simpledialog.askstring("Input", "Ingrese el estado del activo (Activo/Inactivo):").strip()
    #plantilla_mac["User"] = simpledialog.askstring("Input", "Ingrese el nombre del usuario:").strip()
    plantilla_mac["Location"] = simpledialog.askstring("Input", "Ingrese la ubicación del activo:").strip()
    
    # Generar el nombre del activo a partir del usuario
    plantilla_mac["Name"] = f"{plantilla_mac['User']}-MacBookPro"

    return plantilla_mac

def guardar_plantilla_txt(asset_data, nombre_archivo):
    with open(nombre_archivo, 'w') as file:
        for key, value in asset_data.items():
            file.write(f"{key}: {value}\n")
    messagebox.showinfo("Información", f"Plantilla guardada en {nombre_archivo}")

def agregar_a_excel_dell(asset_data):
    try:
        df = pd.read_excel(ruta_excel)
        nuevo_registro = pd.DataFrame([asset_data])  # Convertir la plantilla a un DataFrame de una fila
        df = pd.concat([df, nuevo_registro], ignore_index=True)
        df.to_excel(ruta_excel, index=False)
        messagebox.showinfo("Éxito", "Datos registrados exitosamente en el Excel.")

        # Guardar la plantilla en un archivo .txt
        nombre_archivo_txt = f"C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Templates/{asset_data['Name']}.txt"
        guardar_plantilla_txt(asset_data, nombre_archivo_txt)

    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar los datos: {e}")

def agregar_a_excel_mac(asset_data):
    try:
        if verificar_existencia_en_excel(asset_data["Serial Number"]):
            messagebox.showinfo("Información", f"El activo con serial '{asset_data['Serial Number']}' ya está registrado en el Excel. No se agregará.")
            return
        df = pd.read_excel(ruta_excel)
        nuevo_registro = pd.DataFrame([asset_data])  # Convertir la plantilla a un DataFrame de una fila
        df = pd.concat([df, nuevo_registro], ignore_index=True)
        df.to_excel(ruta_excel, index=False)
        messagebox.showinfo("Éxito", "Datos registrados exitosamente en el Excel.")

        # Guardar la plantilla en un archivo .txt
        nombre_archivo_txt = f"C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Templates/{asset_data['Name']}.txt"
        guardar_plantilla_txt(asset_data, nombre_archivo_txt)

    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar los datos: {e}")
        
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
            messagebox.showinfo("Información", "Laptop Dell detectada. Procesando datos...")
            
            if len(qr_data) > 7:  # Verificar si el QR escaneado contiene más información de la cuenta
                service_tag = extraer_service_tag(qr_data)
            else:
                service_tag = qr_data  # Si solo tiene el serial en QR directamente
            
            if service_tag:
                #messagebox.showinfo("Información", f"Service Tag detectado: {service_tag}")
                confirmacion = simpledialog.askstring("Confirmación", f"¿Es correcto este Service Tag: {service_tag}, desea continuar? (sí/no):").strip().lower()
                if confirmacion not in ["sí", "si", "Si", "Sí"]:
                    messagebox.showinfo("Información", "Operación cancelada por el usuario.")
                    return
                else:
                    if verificar_existencia_en_excel(service_tag):
                        messagebox.showinfo("Información", f"El activo con serial '{service_tag}' ya está registrado en el Excel. No se agregará.")
                        return
                    asset_data = procesar_qr_dell(service_tag)
                    agregar_a_excel_dell(asset_data)
            else:
                messagebox.showerror("Error", "No se detectó un Service Tag válido en el QR escaneado.")
        else:
            messagebox.showerror("Error", "Código QR no corresponde a un equipo Dell.")
    else:
        messagebox.showerror("Error", "No se detectó ningún código QR.")

def manejar_qr_mac():
    qr_data = escanear_qr_con_celular()
    if qr_data:
        if "MacBook" in qr_data or "Serial Number:" in qr_data or qr_data.startswith("C02") or 10 <= len(qr_data) <= 12:  # Detectar si es Mac por patrones comunes
            messagebox.showinfo("Información", "Laptop MacBook detectada. Procesando datos...")

            if len(qr_data) > 12:
                serial_number = extraer_serial_mac(qr_data)
            else:
                serial_number = qr_data

            if serial_number:
                #messagebox.showinfo("Información", f"Serial Number detectado: {serial_number}")
                confirmacion = simpledialog.askstring("Confirmación", f"¿Es correcto este Serial Number: {serial_number} desea continuar? (sí/no):").strip().lower()
                if confirmacion not in ["sí", "si", "Si", "Sí"]:
                    messagebox.showinfo("Información", "Operación cancelada por el usuario.")
                    return
                else:
                    if verificar_existencia_en_excel(serial_number):
                        messagebox.showinfo("Información", f"El activo con serial '{serial_number}' ya está registrado en el Excel. No se agregará.")
                        return
                    asset_data = procesar_qr_mac(serial_number)
                    agregar_a_excel_mac(asset_data)
            else:
                messagebox.showerror("Error", "No se detectó un Serial Number válido en el QR escaneado.")
        else:
            messagebox.showerror("Error", "Código QR no corresponde a un equipo Mac.")
    else:
        messagebox.showerror("Error", "No se detectó ningún código QR.")

def manejar_qr_laptop():
    while True:
        qr_data = escanear_qr_con_celular()

        if qr_data:
            qr_data_lower = qr_data.lower()

            # Detectar si es una laptop Dell
            patrones_dell = ["dell", "service tag", "made in vietnam", "Service tag", "Dell", "S/N", "(S/N)", "SN"]
            if any(keyword in qr_data_lower for keyword in patrones_dell) or qr_data.startswith("CS") or len(qr_data) == 7:
                messagebox.showinfo("Información", "Laptop Dell detectada. Procesando datos...")

                if len(qr_data) > 7:
                    serial_number = extraer_service_tag(qr_data)
                else:
                    serial_number = qr_data  # Si el QR contiene solo el serial

                if serial_number:
                    #messagebox.showinfo("Información", f"Service Tag detectado: {serial_number}")
                    confirmacion = simpledialog.askstring("Confirmación", f"¿Es correcto este Service Tag: {serial_number}, desea continuar? (sí/no): ").strip().lower()
                    if confirmacion in ["sí", "si", "Si", "Sí"]:
                        if verificar_existencia_en_excel(serial_number):
                            messagebox.showinfo("Información", f"El activo con serial '{serial_number}' ya está registrado en el Excel. No se agregará.")
                            return
                        asset_data = procesar_qr_dell(serial_number)
                        agregar_a_excel_dell(asset_data)
                        break
                    else:
                        messagebox.showinfo("Información", "Reintentando escaneo...")
                        continue
                else:
                    messagebox.showerror("Error", "No se detectó un Service Tag válido. Reintentando...")
                    continue

            # Detectar si es una laptop MacBook
            if any(keyword in qr_data_lower for keyword in ["macbook", "serial number"]) or qr_data.startswith("C02") or 10 <= len(qr_data) <= 12:
                messagebox.showinfo("Información", "Laptop MacBook detectada. Procesando datos...")

                if len(qr_data) > 12:
                    serial_number = extraer_serial_mac(qr_data)
                else:
                    serial_number = qr_data

                if serial_number:
                    #messagebox.showinfo("Información", f"Serial Number detectado: {serial_number}")
                    confirmacion = simpledialog.askstring("Confirmación", f"¿Es correcto este Serial Number: {serial_number}, desea continuar? (sí/no): ").strip().lower()
                    if confirmacion in ["sí", "si", "Si", "Sí"]:
                        if verificar_existencia_en_excel(serial_number):
                            messagebox.showinfo("Información", f"El activo con serial '{serial_number}' ya está registrado en el Excel. No se agregará.")
                            return
                        asset_data = procesar_qr_mac(serial_number)
                        agregar_a_excel_mac(asset_data)
                        break
                    else:
                        messagebox.showinfo("Información", "Reintentando escaneo...")
                        continue
                else:
                    messagebox.showerror("Error", "No se detectó un Serial Number válido. Reintentando...")
                    continue

            messagebox.showerror("Error", "Código QR no corresponde a un equipo Dell ni Mac. Reintentando...")
        else:
            messagebox.showerror("Error", "No se detectó ningún código QR. Reintentando...")

def obtener_user_id(session_token, username):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    params = {"searchText": username.strip().lower(), "range": "0-999"}
    response = requests.get(f"{GLPI_URL}/search/User", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        users = response.json().get("data", [])

        for user in users:
            # Manejar valores None de forma segura
            first_name = (user.get('9') or '').strip()
            last_name = (user.get('34') or '').strip()
            username_glpi = (user.get("1") or '').strip()
            
            # Construir el nombre completo
            nombre_completo = f"{first_name} {last_name}".strip()

            # Comparar el nombre normalizado
            if nombre_completo.lower() == username.strip().lower() or username_glpi.lower() == username.strip().lower():
                messagebox.showinfo("Información", f"Usuario encontrado: {nombre_completo}, ID: {user.get('1')}")
                return user.get("1")  # Asegúrate de que '1' es el ID correcto en tu sistema GLPI

        messagebox.showerror("Información", f"No se encontró el usuario '{username}' en GLPI.")
        return None
    else:
        messagebox.showerror("Error", f"Error al buscar el usuario en GLPI: {response.status_code}")
        return None

def actualizar_asset_glpi(session_token, asset_id, asset_data):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    # Obtener el ID del usuario basado en su nombre
    user_id = obtener_user_id(session_token, asset_data["User"])
    messagebox.showinfo("Información", f"ID del usuario encontrado: {user_id}")
    if not user_id:
        messagebox.showerror("Error", f"No se encontró el usuario '{asset_data['User']}' en GLPI.")
        return

    # Determinar el nuevo nombre según el fabricante
    if "Dell" in asset_data["Manufacturer"] or "Dell Inc." in asset_data["Manufacturer"] or "dell" in asset_data["Manufacturer"] or "DELL" in asset_data["Manufacturer"] or "Dell inc." in asset_data["Manufacturer"]:
        new_name = f"{asset_data['User']}-Latitude"
    elif "Apple" in asset_data["Manufacturer"] or "Apple Inc." in asset_data["Manufacturer"] or "apple" in asset_data["Manufacturer"] or "MAC" in asset_data["Manufacturer"] or "Mac" in asset_data["Manufacturer"] or "mac" in asset_data["Manufacturer"]:
        new_name = f"{asset_data['User']}-MacBookPro"
    else:
        messagebox.showerror("Error", "No se pudo determinar el fabricante del laptop.")
        return

    # Preparar datos para la actualización en GLPI
    payload = {
        "input": {
            "id": asset_id,  
            "name": new_name,
            #"users_id": user_id
        }
    }

    response = requests.put(f"{GLPI_URL}/Computer/{asset_id}", headers=headers, json=payload, verify=False)

    if response.status_code == 200:
        messagebox.showinfo("Éxito", f"Activo con ID {asset_id} actualizado correctamente en GLPI con el nombre '{new_name}'.")
    else:
        messagebox.showerror("Error", f"Error al actualizar el activo en GLPI: {response.status_code}")
        try:
            messagebox.showerror("Error", response.json())
        except json.JSONDecodeError:
            messagebox.showerror("Error", response.text)

def obtener_asset_id_por_serial(session_token, serial_number):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    params = {
        "searchText": serial_number.strip().lower(),
        "range": "0-999"
    }

    # Buscar el asset por número de serie
    response = requests.get(f"{GLPI_URL}/search/Computer", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        assets = response.json().get("data", [])
        
        for asset in assets:
            serial_found = (asset.get("5") or "").strip().lower()  # Clave 5 es el serial number
            asset_name = asset.get("1")  # Clave 1 es el nombre del asset
            
            if serial_found == serial_number.lower():
                messagebox.showinfo("Información", f"Activo encontrado: {asset_name}, Serial: {serial_number}")
                # Ahora buscamos el ID utilizando el nombre del activo encontrado
                return obtener_id_por_nombre(session_token, asset_name)

        messagebox.showinfo("Información", f"No se encontró un activo con el serial number '{serial_number}' en GLPI.")
        return None
    else:
        messagebox.showerror("Error", f"Error al buscar el activo en GLPI: {response.status_code}")
        return None

def obtener_id_por_nombre(session_token, asset_name):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    params = {
        "searchText": asset_name.strip().lower(),
        "range": "0-999"
    }

    # Buscar el asset por nombre
    response = requests.get(f"{GLPI_URL}/Computer", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        for asset in response.json():
            if asset.get("name").strip().lower() == asset_name.strip().lower():
                messagebox.showinfo("Información", f"Activo encontrado: {asset_name}, ID: {asset.get('id')}")
                return asset.get("id")

        messagebox.showinfo("Información", f"No se encontró el ID para el activo '{asset_name}' en GLPI.")
        return None
    else:
        messagebox.showerror("Error", f"Error al buscar el ID del activo: {response.status_code}")
        return None

def entregar_laptop():
    try: 
        messagebox.showinfo("Información", "--- Entregar Laptop a Usuario ---")
        manufacturer = simpledialog.askstring("Input", "Ingrese el fabricante del laptop (Dell/Mac):").strip().lower()
        serial_number = None
        if manufacturer == "Dell" or manufacturer == "dell" or manufacturer == "Dell inc." or manufacturer == "Dell Inc." or manufacturer == "DELL":
            metodo = simpledialog.askstring("Input", "¿Desea escanear el QR o ingresar el Service Tag manualmente? (escanear/manual):").strip().lower()
            if metodo == "escanear":
                qr_data = escanear_qr_con_celular()
                if re.match(r'\bcs[a-z0-9]{5}\b', qr_data) or re.match(r'^[A-Za-z0-9]{7}$', qr_data):
                    messagebox.showinfo("Información", "Laptop Dell detectada. Procesando datos...")
                    serial_number = qr_data
            elif metodo == "manual":
                serial_number = simpledialog.askstring("Input", "Ingrese el Service Tag del laptop:").strip()
                if not re.match(r'^[A-Za-z0-9]{7}$', serial_number) or not re.match(r'\bcs[a-z0-9]{5}\b', qr_data):
                    messagebox.showerror("Error", "Service Tag no válido. Intente nuevamente.")
                    return
                else:
                    messagebox.showinfo("Información", "Laptop Dell detectada. Procesando datos...")
                    serial_number = qr_data
            else:
                messagebox.showerror("Error", "Método no válido. Intente nuevamente.")
                return
        elif manufacturer == "Mac" or manufacturer == "mac" or manufacturer == "Mac Inc." or manufacturer == "Apple Inc." or manufacturer == "Apple":
            metodo = simpledialog.askstring("Input", "¿Desea escanear el QR o ingresar el Serial Number manualmente? (escanear/manual):").strip().lower()
            if metodo == "escanear":
                qr_data = escanear_qr_con_celular()
                if re.match(r'\bC02[A-Za-z0-9]{8,10}\b', qr_data) or re.match(r'^[A-Za-z0-9]{10,12}$', qr_data) or re.match(r'^S?[C02][A-Za-z0-9]{8}$', qr_data) or re.match(r'^S[A-Za-z0-9]{9}$', qr_data):
                    messagebox.showinfo("Información", "Laptop MacBook detectada. Procesando datos...")
                    # Remover la 'S' del serial number si existe al inicio
                    serial_number = qr_data
                    serial_number = serial_number[1:] if serial_number.startswith("S") else serial_number
            elif metodo == "manual": 
                serial_number = simpledialog.askstring("Input", "Ingrese el Service Tag del laptop:").strip()
                if not re.match(r'\bC02[A-Za-z0-9]{8,10}\b', qr_data) or not re.match(r'^[A-Za-z0-9]{10,12}$', qr_data) or not re.match(r'^S?[C02][A-Za-z0-9]{8}$', qr_data) or not re.match(r'^S[A-Za-z0-9]{9}$', qr_data):
                    messagebox.showerror("Error", "Service Tag no válido. Intente nuevamente.")
                    return
                else:
                    messagebox.showinfo("Información", "Laptop Mac detectada. Procesando datos...")
                    serial_number = qr_data

        else:
            messagebox.showerror("Error", "Fabricante no válido. Intente nuevamente.")
            return

        # Cargar el archivo Excel
        df = pd.read_excel(ruta_excel)
        if df.empty:
            messagebox.showerror("Error", "El archivo Excel está vacío.")
            return

        # Validar columnas necesarias
        required_columns = ["Serial Number", "Manufacturer", "User", "Name"]
        if not all(col in df.columns for col in required_columns):
            messagebox.showerror("Error", "El archivo Excel no contiene las columnas necesarias.")
            return
        
        filtro = df[df["Serial Number"].str.lower() == serial_number.lower()]

        if filtro.empty:
            messagebox.showerror("Error", f"No se encontró un laptop con el Service Tag '{serial_number}' en el archivo Excel.")
            return

        nuevo_usuario = simpledialog.askstring("Input", "Ingrese el nombre del usuario que recibirá el laptop:").strip()
        if not nuevo_usuario:
            messagebox.showerror("Error", "El nombre del usuario no puede estar vacío.")
            return
        
        # Manejar valores NaN antes de actualizar el DataFrame
        df["User"] = df["User"].fillna("")
        df["Name"] = df["Name"].fillna("Unknown")


        # Determinar el nuevo nombre del laptop en base al fabricante
        if manufacturer == "Dell" or manufacturer == "dell" or manufacturer == "Dell inc." or manufacturer == "Dell Inc." or manufacturer == "DELL":
            new_name = f"{nuevo_usuario}-Latitude"
        elif manufacturer == "Mac" or manufacturer == "mac" or manufacturer == "Mac Inc." or manufacturer == "Apple Inc." or manufacturer == "Apple":
            new_name = f"{nuevo_usuario}-MacBookPro"
        else:
            messagebox.showerror("Error", "No se pudo determinar el fabricante del laptop.")
            return

        # Actualizar DataFrame con los nuevos valores
        df.loc[df["Serial Number"].str.lower() == serial_number.lower(), "User"] = nuevo_usuario
        df.loc[df["Serial Number"].str.lower() == serial_number.lower(), "Name"] = new_name

        df.to_excel(ruta_excel, index=False)
        messagebox.showinfo("Información", f"Laptop con Service Tag '{serial_number}' asignado a '{nuevo_usuario}' en el Excel.")

        # Actualizar en GLPI
        session_token = obtener_token_sesion()
        if not session_token:
            messagebox.showerror("Error", "No se pudo obtener el token de sesión.")
            return

        asset_id = obtener_asset_id_por_serial(session_token, serial_number)
        if not asset_id:
            messagebox.showerror("Error", "No se pudo encontrar el activo en GLPI.")
            return

        asset_data = filtro.iloc[0].to_dict()
        asset_data["User"] = nuevo_usuario
        asset_data["Name"] = new_name  

        actualizar_asset_glpi(session_token, asset_id, asset_data)
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error inesperado: {str(e)}")

def procesar_qr_monitor(qr_data):
    # Plantilla para monitores
    plantilla_monitor = {
        "Asset Type": "Monitor",
        "Status": "Stocked",  # Definir un estado predeterminado
        "User": None,  # Pedir al usuario
        "Name": None,  # Generado automáticamente
        "Location": None,  # Pedir al usuario
        "Manufacturer": "Dell Inc.",  # Extraído del QR
        "Model": None,  # Extraído del QR
        "Serial Number": qr_data.strip(),  # Código QR escaneado
        "Comments": "Check",
    }

    # Solicitar datos adicionales al usuario
    plantilla_monitor["Location"] = simpledialog.askstring("Input", "Ingrese la ubicación del monitor:").strip()
    #plantilla_monitor["Manufacturer"] = simpledialog.askstring("Input", "Ingrese el fabricante del monitor:").strip()
    #plantilla_monitor["Model"] = simpledialog.askstring("Input", "Ingrese el modelo del monitor:").strip()

    # Generar el nombre del activo a partir del modelo
    plantilla_monitor["Name"] = f"{plantilla_monitor['Model']}-{plantilla_monitor['Serial Number']}"

    return plantilla_monitor

def manejar_qr_monitor():
    qr_data = escanear_qr_con_celular()
    if qr_data:
        if any(keyword in qr_data.lower() for keyword in ["monitor", "display", "serial number", "CN-", "SN", "S/N", "CN"]) or len(qr_data) == 7 or len(qr_data) > 12:
            messagebox.showinfo("Información", "Monitor detectado. Procesando datos...")

            serial_number = qr_data.strip()

            if serial_number:
                messagebox.showinfo("Información", f"Serial Number detectado: {serial_number}")
                confirmacion = simpledialog.askstring("Confirmación", "¿Es correcto este Serial Number, desea continuar? (sí/no): ").strip().lower()
                if confirmacion not in ["sí", "si"]:
                    messagebox.showinfo("Información", "Operación cancelada por el usuario.")
                    return
                else:
                    if verificar_existencia_en_excel(serial_number):
                        messagebox.showinfo("Información", f"El activo con serial '{serial_number}' ya está registrado en el Excel. No se agregará.")
                        return
                    asset_data = procesar_qr_monitor(serial_number)
                    agregar_a_excel(asset_data)
        else:
            messagebox.showerror("Error", "Código QR no corresponde a un monitor.")
    else:
        messagebox.showerror("Error", "No se detectó ningún código QR.")

def actualizar_asset_glpi_monitor(session_token, asset_id, asset_data):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    # Obtener el ID del usuario basado en su nombre
    user_id = obtener_user_id(session_token, asset_data["User"])
    messagebox.showinfo("Información", f"ID del usuario encontrado: {user_id}")
    if not user_id:
        messagebox.showerror("Error", f"No se encontró el usuario '{asset_data['User']}' en GLPI.")
        return

    # Determinar el nuevo nombre del monitor
    if "Dell" in asset_data["Manufacturer"]:
        new_name = f"{asset_data['User']}-DellMonitor"
    elif "Samsung" in asset_data["Manufacturer"]:
        new_name = f"{asset_data['User']}-SamsungMonitor"
    else:
        new_name = f"{asset_data['User']}-Monitor"

    # Preparar datos para la actualización en GLPI
    payload = {
        "input": {
            "id": asset_id,  
            "name": new_name,
            "users_id": user_id
        }
    }

    response = requests.put(f"{GLPI_URL}/Monitor/{asset_id}", headers=headers, json=payload, verify=False)

    if response.status_code == 200:
        messagebox.showinfo("Éxito", f"Monitor con ID {asset_id} actualizado correctamente en GLPI con el nombre '{new_name}'.")
    else:
        messagebox.showerror("Error", f"Error al actualizar el monitor en GLPI: {response.status_code}")
        try:
            messagebox.showerror("Error", response.json())
        except json.JSONDecodeError:
            messagebox.showerror("Error", response.text)

def obtener_id_por_nombre_monitor(session_token, asset_name):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    params = {
        "searchText": asset_name.strip().lower(),
        "range": "0-999"
    }

    response = requests.get(f"{GLPI_URL}/Monitor", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        for asset in response.json():
            if asset.get("name").strip().lower() == asset_name.strip().lower():
                messagebox.showinfo("Información", f"Monitor encontrado: {asset_name}, ID: {asset.get('id')}")
                return asset.get("id")

        messagebox.showinfo("Información", f"No se encontró el ID para el monitor '{asset_name}' en GLPI.")
        return None
    else:
        messagebox.showerror("Error", f"Error al buscar el ID del monitor: {response.status_code}")
        return None

def obtener_asset_id_por_serial_monitor(session_token, serial_number):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    params = {
        "searchText": serial_number.strip().lower(),
        "range": "0-999"
    }

    response = requests.get(f"{GLPI_URL}/search/Monitor", headers=headers, params=params, verify=False)

    if response.status_code == 200:
        assets = response.json().get("data", [])
        
        for asset in assets:
            serial_found = (asset.get("5") or "").strip().lower()  # Clave 5 es el serial number
            asset_name = asset.get("1")  # Clave 1 es el nombre del asset
            
            if serial_found == serial_number.lower():
                messagebox.showinfo("Información", f"Monitor encontrado: {asset_name}, Serial: {serial_number}")
                return obtener_id_por_nombre_monitor(session_token, asset_name)

        messagebox.showinfo("Información", f"No se encontró un monitor con el número de serie '{serial_number}' en GLPI.")
        return None
    else:
        messagebox.showerror("Error", f"Error al buscar el monitor en GLPI: {response.status_code}")
        return None

def entregar_monitor():
    messagebox.showinfo("Información", "--- Entregar Monitor a Usuario ---")
    metodo = simpledialog.askstring("Input", "¿Desea escanear el QR o ingresar el número de serie manualmente? (escanear/manual):").strip().lower()

    if metodo == "escanear":
        qr_data = escanear_qr_con_celular()
        if any(keyword in qr_data.lower() for keyword in ["monitor", "serial number", "sn", "cn"]) or len(qr_data) >= 7:
            messagebox.showinfo("Información", "Monitor detectado. Procesando datos...")
            serial_number = qr_data.strip()
        else:
            messagebox.showerror("Error", "Código QR no corresponde a un monitor válido.")
            return
    elif metodo == "manual":
        serial_number = simpledialog.askstring("Input", "Ingrese el número de serie del monitor:").strip()
    else:
        messagebox.showerror("Error", "Método no válido. Intente nuevamente.")
        return

    df = pd.read_excel(ruta_excel)
    if df.empty:
        messagebox.showerror("Error", "El archivo Excel está vacío.")
        return

    filtro = df[df["Serial Number"].str.lower() == serial_number.lower()]

    if filtro.empty:
        messagebox.showerror("Error", f"No se encontró un monitor con el número de serie '{serial_number}' en el archivo Excel.")
        return

    nuevo_usuario = simpledialog.askstring("Input", "Ingrese el nombre del usuario que recibirá el monitor:").strip()

    # Manejar valores NaN antes de actualizar el DataFrame
    df["User"] = df["User"].fillna("")
    df["Name"] = df["Name"].fillna("Unknown")

    # Determinar el nuevo nombre del monitor en base al usuario
    fabricante = filtro["Manufacturer"].values[0]
    if "Dell" in fabricante:
        new_name = f"{nuevo_usuario}-DellMonitor"
    elif "Samsung" in fabricante:
        new_name = f"{nuevo_usuario}-SamsungMonitor"
    else:
        new_name = f"{nuevo_usuario}-Monitor"

    df.loc[df["Serial Number"].str.lower() == serial_number.lower(), "User"] = nuevo_usuario
    df.loc[df["Serial Number"].str.lower() == serial_number.lower(), "Name"] = new_name

    df.to_excel(ruta_excel, index=False)
    messagebox.showinfo("Información", f"Monitor con número de serie '{serial_number}' asignado a '{nuevo_usuario}' en el Excel.")

    # Actualizar en GLPI
    session_token = obtener_token_sesion()
    if not session_token:
        messagebox.showerror("Error", "No se pudo obtener el token de sesión.")
        return

    asset_id = obtener_asset_id_por_serial_monitor(session_token, serial_number)
    if not asset_id:
        messagebox.showerror("Error", "No se pudo encontrar el activo en GLPI.")
        return

    asset_data = filtro.iloc[0].to_dict()
    asset_data["User"] = nuevo_usuario
    asset_data["Name"] = new_name  

    actualizar_asset_glpi_monitor(session_token, asset_id, asset_data)

def agregar_consumible():
    messagebox.showinfo("Información", "--- Agregar Consumible al Stock ---")

    metodo = simpledialog.askstring("Input", "¿Desea escanear el QR o ingresar manualmente? (qr/manual): ").strip().lower()

    if metodo == "qr":
        qr_data = escanear_qr_con_celular()
        if qr_data:
            inventory_number = qr_data.strip()
            messagebox.showinfo("Información", f"Inventory Number detectado: {inventory_number}")
        else:
            messagebox.showerror("Error", "No se detectó ningún código QR.")
            return
    else:
        inventory_number = simpledialog.askstring("Input", "Ingrese el número de inventario o activo: ").strip()

    df = pd.read_excel(ruta_excel_consumibles)
    df.columns = df.columns.str.strip()  # Asegura que no haya espacios en los nombres de columnas

    # Verificar si el consumible ya está registrado en el Excel
    filtro = df[df["Inventory/Asset Number"].astype(str).str.lower() == inventory_number.lower()]

    if not filtro.empty:
        messagebox.showinfo("Información", f"Consumible con Inventory Number '{inventory_number}' encontrado en el Excel.")
        nombre_consumible = filtro.iloc[0]["Name"]
        location = filtro.iloc[0]["Location"]
    else:
        messagebox.showinfo("Información", f"No se encontró un consumible con el número de inventario '{inventory_number}'. Creando nuevo...")
        nombre_consumible = simpledialog.askstring("Input", "Ingrese el nombre del nuevo consumible: ").strip()
        location = simpledialog.askstring("Input", "Ingrese la ubicación del consumible: ").strip()

    cantidad = int(simpledialog.askstring("Input", "Ingrese la cantidad a agregar al stock: "))

    session_token = obtener_token_sesion()
    if not session_token:
        messagebox.showerror("Error", "No se pudo obtener el token de sesión.")
        return

    # Corrección: Se pasa tanto el nombre como el inventory_number a la función
    consumible_id = obtener_id_consumible(session_token, nombre_consumible, inventory_number)
    
    if not consumible_id:
        messagebox.showinfo("Información", f"No se encontró el consumible '{nombre_consumible}' en GLPI. Creando uno nuevo...")
        consumible_id = crear_consumible(session_token, nombre_consumible, inventory_number, location, cantidad)
        if not consumible_id:
            messagebox.showerror("Error", "Error al crear el consumible en GLPI.")
            return

    stock_actual = obtener_stock_actual(session_token, consumible_id)
    nuevo_stock = stock_actual + cantidad

    actualizar_stock_glpi(session_token, consumible_id, nuevo_stock)
    messagebox.showinfo("Información", f"Consumible '{nombre_consumible}' actualizado a {nuevo_stock} unidades en stock.")

    # Registrar en Excel
    actualizar_excel_consumible(nombre_consumible, inventory_number, location, nuevo_stock)

def quitar_consumible():
    messagebox.showinfo("Información", "--- Quitar Consumible del Stock ---")

    metodo = simpledialog.askstring("Input", "¿Desea escanear el QR o ingresar manualmente? (qr/manual): ").strip().lower()

    if metodo == "qr":
        qr_data = escanear_qr_con_celular()
        if qr_data:
            inventory_number = qr_data.strip()
            messagebox.showinfo("Información", f"Inventory Number detectado: {inventory_number}")
        else:
            messagebox.showerror("Error", "No se detectó ningún código QR.")
            return
    else:
        inventory_number = simpledialog.askstring("Input", "Ingrese el número de inventario o activo: ").strip()

    df = pd.read_excel(ruta_excel_consumibles)
    df.columns = df.columns.str.strip()  # Asegurar que no haya espacios en los nombres de columnas

    # Buscar el consumible en el Excel por inventory_number
    filtro = df[df["Inventory/Asset Number"].astype(str).str.lower() == inventory_number.lower()]

    if filtro.empty:
        messagebox.showerror("Error", f"No se encontró el consumible con Inventory Number '{inventory_number}' en el archivo Excel.")
        return

    # Obtener datos existentes del consumible
    nombre_consumible = filtro.iloc[0]["Name"]
    location = filtro.iloc[0]["Location"]
    stock_actual = int(filtro.iloc[0]["Stock Target"])

    cantidad = int(simpledialog.askstring("Input", f"Ingrese la cantidad a retirar (Stock actual: {stock_actual}): "))

    if stock_actual < cantidad:
        messagebox.showerror("Error", "No se puede retirar más cantidad de la que hay en stock.")
        return

    nuevo_stock = stock_actual - cantidad

    session_token = obtener_token_sesion()
    if not session_token:
        messagebox.showerror("Error", "No se pudo obtener el token de sesión.")
        return

    consumible_id = obtener_id_consumible(session_token, nombre_consumible, inventory_number)
    if not consumible_id:
        messagebox.showerror("Error", f"No se encontró el consumible '{nombre_consumible}' en GLPI.")
        return

    actualizar_stock_glpi(session_token, consumible_id, nuevo_stock)
    messagebox.showinfo("Información", f"Consumible '{nombre_consumible}' actualizado a {nuevo_stock} unidades en stock.")

    # Actualizar en Excel
    actualizar_excel_consumible(nombre_consumible, inventory_number, location, nuevo_stock)

def actualizar_excel_consumible(nombre, inventory_number, location, stock):
    df = pd.read_excel(ruta_excel_consumibles)
    df.columns = df.columns.str.strip()  # Asegura nombres sin espacios extra

    # Convertir inventory_number a string para evitar errores
    inventory_number_str = str(inventory_number).strip().lower()

    # Verificar si el consumible ya está registrado
    filtro = df[
        (df["Name"].str.lower() == nombre.lower()) & 
        (df["Inventory/Asset Number"].astype(str).str.lower() == inventory_number_str)
    ]

    if not filtro.empty:
        df.loc[
            (df["Name"].str.lower() == nombre.lower()) &
            (df["Inventory/Asset Number"].astype(str).str.lower() == inventory_number_str), 
            "Stock Target"
        ] = stock
    else:
        nuevo_consumible = pd.DataFrame([{
            "Name": nombre,
            "Inventory/Asset Number": inventory_number_str,  # Convertir a string
            "Location": location,
            "Stock Target": stock
        }])
        df = pd.concat([df, nuevo_consumible], ignore_index=True)

    df.to_excel(ruta_excel_consumibles, index=False)
    messagebox.showinfo("Información", f"El consumible '{nombre}' ha sido registrado/actualizado en el Excel.")

def crear_consumible(session_token, nombre, inventory_number, location, stock_target):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    payload = {
        "input": {
            "name": nombre,
            "otherserial": inventory_number,
            "locations_id": obtener_location_id(session_token, location),
            "stock_target": stock_target
        }
    }

    response = requests.post(f"{GLPI_URL}/ConsumableItem", headers=headers, json=payload, verify=False)

    if response.status_code == 201:
        consumible_id = response.json().get("id")
        messagebox.showinfo("Información", f"Consumible '{nombre}' creado exitosamente con ID {consumible_id}.")
        return consumible_id
    else:
        messagebox.showerror("Error", f"Error al crear el consumible: {response.status_code}")
        try:
            messagebox.showerror("Error", response.json())
        except json.JSONDecodeError:
            messagebox.showerror("Error", response.text)
        return None

def obtener_id_consumible(session_token, nombre_consumible, inventory_number):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    params = {
        "searchText": nombre_consumible.strip().lower(),
        "range": "0-999"
    }

    response = requests.get(f"{GLPI_URL}/ConsumableItem", headers=headers, params=params, verify=False)
    
    if response.status_code == 200:
        consumibles = response.json()
        for consumible in consumibles:
            # Convertir a cadena de texto y limpiar espacios en blanco
            consumible_name = str(consumible.get("name", "")).strip().lower()
            consumible_serial = str(consumible.get("otherserial", "")).strip().lower()
            inventory_number_str = str(inventory_number).strip().lower()

            if consumible_name == nombre_consumible.strip().lower() and consumible_serial == inventory_number_str:
                return consumible["id"]

    messagebox.showinfo("Información", f"No se encontró el consumible con nombre '{nombre_consumible}' y número de inventario '{inventory_number}' en GLPI.")
    return None

def obtener_stock_actual(session_token, consumible_id):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    response = requests.get(f"{GLPI_URL}/ConsumableItem/{consumible_id}", headers=headers, verify=False)

    if response.status_code == 200:
        return int(response.json().get("stock_target", 0))
    else:
        messagebox.showerror("Error", f"Error al obtener el stock del consumible ID {consumible_id}: {response.status_code}")
        return 0

def actualizar_stock_glpi(session_token, consumible_id, nuevo_stock):
    headers = {
        "Content-Type": "application/json",
        "Session-Token": session_token,
        "App-Token": APP_TOKEN
    }

    payload = {
        "input": {
            "id": consumible_id,
            "stock_target": nuevo_stock
        }
    }

    response = requests.put(f"{GLPI_URL}/ConsumableItem/{consumible_id}", headers=headers, json=payload, verify=False)

    if response.status_code == 200:
        messagebox.showinfo("Información", f"Stock del consumible ID {consumible_id} actualizado correctamente a {nuevo_stock}.")
    else:
        messagebox.showerror("Error", f"Error al actualizar el stock en GLPI: {response.status_code}")
        try:
            messagebox.showerror("Error", response.json())
        except json.JSONDecodeError:
            messagebox.showerror("Error", response.text)

def salir():
    root.destroy()

root = tk.Tk()
root.title("GLPI Asset Automator")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

tk.Label(frame, text="----- Laptops -----").pack()
tk.Button(frame, text="Escanear QR y registrar cualquier laptop (Dell/Mac), ¡Me siento con suerte!", command=manejar_qr_laptop).pack()
tk.Button(frame, text="Escanear QR y registrar en Excel (Template Default)", command=lambda: agregar_a_excel(parse_qr_data(escanear_qr_con_celular()))).pack()
tk.Button(frame, text="Escanear QR y registrar laptops Dell", command=manejar_qr_dell).pack()
tk.Button(frame, text="Escanear QR y registrar laptops Mac", command=manejar_qr_mac).pack()
tk.Button(frame, text="Entregar laptop a un usuario", command=entregar_laptop).pack()

tk.Label(frame, text="----- Monitores -----").pack()
tk.Button(frame, text="Escanear QR y registrar monitores", command=manejar_qr_monitor).pack()
tk.Button(frame, text="Entregar monitor a un usuario", command=entregar_monitor).pack()

tk.Label(frame, text="----- Consumibles -----").pack()
tk.Button(frame, text="Agregar consumible", command=agregar_consumible).pack()
tk.Button(frame, text="Quitar consumible", command=quitar_consumible).pack()

tk.Label(frame, text="----- Excel -----").pack()
tk.Button(frame, text="Registrar la última fila del Excel en GLPI", command=registrar_ultima_fila).pack()
tk.Button(frame, text="Registrar un activo por nombre", command=registrar_por_nombre).pack()
tk.Button(frame, text="Registrar todos los activos de Excel en GLPI", command=lambda: procesar_archivo_excel(ruta_excel)).pack()
tk.Button(frame, text="Salir", command=salir).pack()

root.mainloop()