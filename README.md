# Descripcion del proyecto

**Práctica:
Automatización de la Gestión de Activos y Consumibles en GLPI**

* **Objetivo de la Práctica:** Actualmente, el grupo de ITG maneja una gran cantidad de activos, incluyendo laptops, monitores y consumibles, que deben ser registrados y gestionados en el sistema GLPI (Gestionnaire Libre de Parc Informatique). Para mejorar la eficiencia y precisión en la gestión de estos activos, se requiere desarrollar una aplicación de escritorio que automatice el proceso de escaneo, registro y actualización de activos y consumibles en GLPI, así como en archivos Excel. La aplicación debe proporcionar una interfaz gráfica de usuario (GUI) intuitiva que permita a los usuEarios realizar estas tareas de manera sencilla y eficiente.

**Entregables:**

- **Automatización del Registro de Activos:** La aplicación debe permitir el escaneo de códigos QR de laptops Dell y Mac, así como de monitores, y registrar estos activos en un archivo Excel y en GLPI. Esto incluye la extracción de información relevante como el número de serie y la generación automática de nombres de activos basados en el usuario y el fabricante.
- **Gestión de Inventario de Consumibles:** La aplicación debe permitir agregar y retirar consumibles del inventario, ya sea escaneando un código QR o ingresando manualmente el número de inventario. El inventario debe mantenerse actualizado en tiempo real tanto en el archivo Excel como en GLPI.
- **Interfaz de Usuario Intuitiva:** La aplicación debe proporcionar una interfaz gráfica fácil de usar, construida con Tkinter, que permita a los usuarios realizar tareas complejas de gestión de activos sin necesidad de conocimientos técnicos avanzados. Los botones deben estar organizados en secciones (Laptops, Monitores, Consumibles, Excel) y cada botón debe estar vinculado a una función específica.
- **Integración con GLPI:** La aplicación debe integrar directamente con GLPI, permitiendo la sincronización de datos entre el archivo Excel y el sistema de gestión de activos. Esto asegura que la información esté siempre actualizada y consistente.
- **Registro en GLPI:** La aplicación debe ofrecer opciones para registrar activos en GLPI directamente desde el archivo Excel, ya sea registrando la última fila, buscando por nombre, o procesando todo el archivo.
- **Código en Repositorio y Manual de Usuario:** El código de la aplicación debe estar disponible en un repositorio, junto con un manual de usuario que explique cómo utilizar la aplicación y cómo configurar las variables de entorno necesarias.
- **Presentación Final:** Se debe preparar una presentación en PowerPoint para el grupo de TI, explicando las funcionalidades de la aplicación, los beneficios de su uso y los resultados obtenidos durante la práctica.

En resumen,esta práctica busca desarrollar una solución que simplifique y automatice la gestión de activos y consumibles en una organización, mejorando la eficiencia operativa y reduciendo los errores humanos. La aplicación resultante permitirá una gestión más precisa y rápida de los activos, asegurando que la información
esté siempre actualizada y disponible para el equipo de ITG.

# Instalacion de glpi en local a traves de Docker

Utilizar el docker-compose.yml el cual levanta todos los servicios necesarios para ejecutar glpi de forma local

```
docker-compose up -d
```

# GLPI conf. inicial

Réplicas SQL (MariaDB o MySQL): db

Usuario SQL: glpi_user

Password SQL: glpi_password

Luego elegir la base de datos existente: glpi

![1737324268611](image/README/1737324268611.png)

Usuario: glpi

Contraseña: glpi

# Estructura de carpetas

- tests/: Legacy code, que almacena versiones o funcionalidad anteriores
- image/: imagenes varias
- Templates/: Archivos de texto para generar el qr de ALMA GLPI
- root/:
  - app.py: ejecuta el codigo principal.
  - .env.example: se utilizan variables de entornos para realizar el procedimiento, y este es un ejemplo de que se debe llenar.
  - docker-compose.yml: Levanta todos los servicios necesarios para poder utilizar GLPI de forma local.
  - Excel-tests.xlsx: Archivo excel que contiene la estructura escencial de los qr's de activos.
  - consumibles.xlsx: Archivo excel que contiene la estructura escencial de los qr's de consumibles.

# Funcionamiento del codigo

# Crear ejecutable

Para crear el ejecutable basta con el comando:

```
pyinstaller --onefile --add-data ".env;." --add-data "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Excel-tests.xlsx;." --add-data "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/consumibles.xlsx;." interface.py
```

En caso de errores, relacionados a pyzbar, basta con agregar:

libiconv.dll

libzbar-64.dll

Y luego ejecutar:

```
pyinstaller --onefile --add-data ".env;." --add-data "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Excel-tests.xlsx;." --add-data "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/consumibles.xlsx;." --add-binary "libiconv.dll;." interface.py
```

```
pyinstaller --onefile --add-data ".env;." --add-data "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/Excel-tests.xlsx;." --add-data "C:/Users/sebastian.salgado/Desktop/GLPI-Asset-Automator/consumibles.xlsx;." --add-binary "libzbar-64.dll;." interface.py
```
