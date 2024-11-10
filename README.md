# Proyecto de Análisis BESI XLXS

Este proyecto es un programa en Python que carga 4 archivos en formato `.xls`, analiza los datos y genera diversos reportes. El objetivo es automatizar la lectura y análisis de datos para generar reportes en un entorno de producción.

## Funcionaliodades

- Carga de archivos: Importa 2 archivos .xls para su análisis.

- Anpalisi8s de datos: Extrae información clave y realiza cálculos automáticos.

- Generación de reportes: Genera reportes automátizados en formato XLSX.

- Exportación: Permite descargar reportes en formato .xlsx y .pdf con un clic.

## Estructura del proyecto

```plaintext
besireports/
├── assets/
│   ├── ico.ico            # Ícono para el ejecutable
│   └── logo.png           # Logo para el reporte .pdf
├── src/
│   ├── data/              # Constantes del programa
│   ├── utils/             # Funciones helper
│   ├── gui.py             # Creación de la interfaz
│   └── main.py            # Entrada al código principal del programa
├── example-docs/          # Archivos de ejemplo para pruebas
│   ├── besi.xlsx
│   └── bom.xlsx
├── README.md              # Documentación del proyecto
├── LICENCE.txt            # Licencia del proyecto
└── requirements.txt       # Requerimientos del proyecto

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
    git clone https://github.com/SebasOrtiz04/besireports.git
    cd besireport

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

4. Puedes descargar el reporte completo en formato XLXS presionando el botón de DESCARGAR.xlxs.

## Build para Distribución con PyInstaller

1. Crea el ejecutable:

    '''bash
    pyinstaller --onedir --windowed --icon=./assets/ico.ico  --add-data "assets\*;assets" src/main.py

2. Arranca el ejecutable generado en /dist/main

## Licencia

Este proyecto está licenciado bajo los términos de la Licencia MIT. Consulta el archivo LICENSE para más detalles.
