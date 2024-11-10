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

## Instalación

1. Clona el repositorio:

    ```bash
    git clone https://github.com/tuusuario/nombre-del-repositorio.git
    cd nombre-del-repositorio

2. Instala las dependencias necesarias:

    pip install -r requirements.txt

## Uso

1. Para ejecutar el programa y generar los reportes, ejecuta:

    python main.py

2. Seleciona los archivos con los encabezados correspondientes:

    - besi.xlxs
    - bom.xlxs

    Podrás encontrar los archivos de ejemplo en la carpeta /example-docs.

3. Selecciona la carpeta para guardar el reporte de surtimiento.

4. Puedes descargar el reporte completo en formato XLXS presionando el botón de 

Este `README.md` cubre el propósito y el uso del programa, cómo configurarlo y los pasos para el despliegue y la generación de un ejecutable en producción. Asegúrate de personalizarlo de acuerdo con las necesidades específicas de tu proyecto.
