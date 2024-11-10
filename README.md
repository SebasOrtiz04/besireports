# Análisis BESI

# Proyecto de Análisis BESI XLXS

Este proyecto es un programa en Python que carga 4 archivos en formato `.xls`, analiza los datos y genera diversos reportes. El objetivo es automatizar la lectura y análisis de datos para generar reportes en un entorno de producción.

## Requisitos

1. Python 3.9 o superior
2. Librerías de Python:
   - `pandas`
   - `numpy`
   - `openpyxl`
   - `pillow`
   - `report-lab`
   - `pyinstaller`

Aquí tienes un ejemplo de un README.md para tu programa. Este archivo incluye una descripción del proyecto, instrucciones para la instalación, el uso, cómo hacer el deploy en producción utilizando watchdog para monitorear cambios en los archivos, y el proceso para hacer un build con pyinstaller.

markdown
Copiar código
## Instalación

1. Clona el repositorio:

   ```bash
   git clone https://github.com/tuusuario/nombre-del-repositorio.git
   cd nombre-del-repositorio
Instala las dependencias necesarias:

bash
Copiar código
pip install -r requirements.txt
El archivo requirements.txt debe incluir las dependencias necesarias:

txt
Copiar código
pandas
watchdog
Asegúrate de tener los archivos .xls en la carpeta especificada en el archivo de configuración o como se indica en las instrucciones de uso.

Uso
Para ejecutar el programa y generar los reportes, ejecuta:

bash
Copiar código
python main.py
El programa cargará los archivos .xls, realizará el análisis y generará cuatro reportes en la carpeta de salida especificada.

Configuración
En el archivo config.json, puedes definir:

La carpeta de entrada donde se encuentran los archivos .xls.
La carpeta de salida para los reportes generados.
Otros parámetros de configuración necesarios para el análisis de datos.
Ejemplo de config.json:

json
Copiar código
{
    "input_folder": "./data/input",
    "output_folder": "./data/output"
}
Deploy en Producción

## Instalación

1. Clona el repositorio:

   ```bash
   git clone https://github.com/tuusuario/nombre-del-repositorio.git
   cd nombre-del-repositorio



Este `README.md` cubre el propósito y el uso del programa, cómo configurarlo y los pasos para el despliegue y la generación de un ejecutable en producción. Asegúrate de personalizarlo de acuerdo con las necesidades específicas de tu proyecto.
