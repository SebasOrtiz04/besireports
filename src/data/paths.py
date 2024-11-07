import os
import sys

# Detectar si estamos en el entorno de ejecución de PyInstaller
if hasattr(sys, "_MEIPASS"):
    # Ruta para cuando la aplicación está empaquetada con PyInstaller
    BASE_PATH = sys._MEIPASS
else:
    # Ruta de desarrollo (raíz del proyecto)
    BASE_PATH = "."

# Rutas de imágenes y archivos
ICON_PATH = os.path.join(BASE_PATH, 'assets', 'motherson.ico')
IMAGE_PATH = os.path.join(BASE_PATH, 'assets', 'logo.png')
FILES_PATH = os.path.join(BASE_PATH, 'docs')
REPORT_NAME = 'Reporte Surtido de Cajas Motherson.xlsx'
PDF_REPORT_NAME = 'Reporte de requerimientos Motherson.pdf'
REPORT_NAME_2 = 'Reporte DOH Motherson.xlsx'
REPORT_NAME_3 = 'Análisis DOH Motherson.xlsx'
