import pandas as pd
import cv2  # Para la captura de QR
from pyzbar.pyzbar import decode  # Decodificar QR
import os
from openpyxl import load_workbook
import np

# Ruta del archivo Excel
ruta_excel = "C:/Users/sebas/Desktop/GLPI-Asset-Automator/Inventario Rittal_SCO y OSF_16 Nov 2021.xlsx"

# Crear archivo Excel si no existe
if not os.path.exists(ruta_excel):
    columnas_necesarias = ["Código", "Componente", "Marca", "Ubicación", "Comentarios"]
    df = pd.DataFrame(columns=columnas_necesarias)
    df.to_excel(ruta_excel, index=False)

# Función para escanear QR usando la cámara
def escanear_qr():
    cap = cv2.VideoCapture(0)
    print("Apunta la cámara al código QR. Presiona 'q' para salir.")

    while True:
        ret, frame = cap.read()
        if not ret:
            print("No se pudo acceder a la cámara.")
            break

        # Convierte a escala de grises (mejora la detección)
        gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # Decodifica los códigos QR en el frame
        qr_codes = decode(gray_frame)

        for qr in qr_codes:
            # Extrae el contenido del QR
            qr_data = qr.data.decode('utf-8')
            print(f"Código QR escaneado: {qr_data}")

            # Resalta el QR en la imagen
            points = qr.polygon
            if len(points) == 4:
                pts = [(point.x, point.y) for point in points]
                pts = cv2.polylines(frame, [np.array(pts, np.int32)], True, (0, 255, 0), 3)

            # Muestra el contenido del QR
            cap.release()
            cv2.destroyAllWindows()
            return qr_data

        # Muestra el frame en una ventana
        cv2.imshow("Escaneando QR", frame)

        # Salir con 'q'
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    return None

# Función para buscar datos asociados al QR o número de serie
def buscar_datos(codigo):
    # Simulando una base de datos local
    base_datos = {
        "12345": {"Componente": "Router Cisco", "Marca": "Cisco", "Ubicación": "Rack A1", "Comentarios": "N/A"},
        "67890": {"Componente": "Switch HP", "Marca": "HP", "Ubicación": "Rack B2", "Comentarios": "Verificado"},
    }
    return base_datos.get(codigo, None)

# Función para agregar datos al Excel
def agregar_a_excel(dato):
    try:
        workbook = load_workbook(ruta_excel)
        sheet = workbook.active
        nueva_fila = [dato["Código"], dato["Componente"], dato["Marca"], dato["Ubicación"], dato["Comentarios"]]
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

    # Buscar los datos asociados
    datos = buscar_datos(codigo)
    if datos:
        datos["Código"] = codigo  # Añadir el código al registro
        agregar_a_excel(datos)
    else:
        print("No se encontraron datos asociados a ese código.")

# Ejecutar la función principal
if __name__ == "__main__":
    main()
