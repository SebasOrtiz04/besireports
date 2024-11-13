import os
import sys

# Detectar si estamos en el entorno de ejecución de PyInstaller
if hasattr(sys, "_MEIPASS"):
    # Ruta para cuando la aplicación está empaquetada con PyInstaller
    BASE_PATH = sys._MEIPASS
else:
    # Ruta de desarrollo (raíz del proyecto)
    BASE_PATH = "./src"

# Rutas de imágenes y archivos
ICON_PATH = os.path.join(BASE_PATH, 'assets', 'ico.ico')
IMAGE_PATH = os.path.join(BASE_PATH, 'assets', 'logo.png')
FILES_PATH = os.path.join(BASE_PATH, 'docs')
REPORT_NAME = 'Reporte Surtido de Cajas Motherson.xlsx'
PDF_REPORT_NAME = 'Reporte de requerimientos Motherson.pdf'
REPORT_NAME_2 = 'Reporte DOH Motherson.xlsx'
REPORT_NAME_3 = 'Análisis DOH Motherson.xlsx'

# Códigos de plataformas
platforms = {
    "Jetta":{
        "label":"Jetta",
        "code":"17A"
    },
    "Tiguan":{
        "label":"Tiguan",
        "code":"55N"
    },
    "Taos":{
        "label":"Taos",
        "code":"2GM"
    },
    "Tayron":{
        "label":"Tayron",
        "code":"57N"
    },
}