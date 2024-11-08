import pandas as pd
import re
import time
import math
import os

from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from data.index import platforms
from data.paths import REPORT_NAME, IMAGE_PATH,REPORT_NAME_3

from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image, PageBreak
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from tkinter import filedialog, messagebox
from datetime import datetime

#-------------------------------------------------------------
def createProgressBar(root, contenedor_botones):
    
    # Configurar el estilo personalizado para la barra de progreso
    style = ttk.Style(root)  
    style.configure("my.Horizontal.TProgressbar", 
                    troughcolor='white',  
                    background='red',    
                    thickness=10) 
    
    # Crear la barra de progreso en modo determinate
    progress_bar = ttk.Progressbar(contenedor_botones, 
                                   style="my.Horizontal.TProgressbar", 
                                   mode='determinate', 
                                   length=200)
    progress_bar.pack(side="left", padx=10)
    progress_bar['value'] = 0  # Inicializa la barra a 0%
    root.update()
    
    return progress_bar

#--------------------------------------------------------------------
def finish_destroy_progress(root, progress_bar):
    progress_bar['value'] = 100
    root.update()
    time.sleep(0.03)
    progress_bar.destroy()
    root.update()
    time.sleep(0.03)

#--------------------------------------------------------------
def getDateHeaders(df):
    
    dateHeaders = []

    # Definir una expresión regular para el formato de fecha DD/MM/YYYY
    fecha_regex = r'^\d{2}/\d{2}/\d{4}'

    # Iterar sobre cada columna y verificar el nombre
    for column in df.columns:
        if re.match(fecha_regex, column):
            dateHeaders.append(column)

    # Mostrar los encabezados que tienen formato de fecha
    return dateHeaders

#---------------------------------------------------------------
def cookBesi(df):

    #Insertar comulnas a calcular
    df.insert(3, 'Referencia', [None] * len(df))
    df.insert(10, 'Dr', [None] * len(df))

    #Extraer los encabezados con formato de fecha
    dateHeaders = getDateHeaders(df)

    # Iterar sobre las filas del DataFrame
    for index, row in df.iterrows():
        # Asignar un valor a la columna 'Referencia' usando .at
        df.at[index, 'Referencia'] = f"{row['TME']}-{row['Noparte']}"

        dalyRate = max(row[dateHeaders])
        df.at[index,'Dr'] = dalyRate
    
    df = df[df['Dr'] > 0 ]
    return df

#---------------------------------------------------------------

def platformFilter(df):

    platformKeys = []
    for key in platforms.keys():
        platformKeys.append(platforms[key]['code'])

    df = df[df['TME'].isin(platformKeys)]

    return df

#--------------------------------------------------------------

def uploadBesi(root, contenedor_botones):
    # Encabezados esperados
    expected_headers = ['TME', 'Noparte']

    # Seleccionar archivo
    archivo = filedialog.askopenfilename(title='Subir BESI',filetypes=[("Archivos Excel", "*.xls;*.xlsx"), ("Todos los archivos", "*.*")])    

    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)
    
    progress_bar['value'] = 0  # Inicializa la barra a 0%
    root.update()

    if archivo:
        try:
            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 10  # 25%
            root.update()
            time.sleep(0.03)
            
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(archivo, header=0)

            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 20  # 25%
            root.update()
            time.sleep(0.03)

            # Asegurar que los encabezados sean strings y eliminar espacios en blanco
            df.columns = df.columns.map(lambda x: str(x).strip())

            if set(expected_headers).issubset(df.columns):

                progress_bar['value'] = 40  # 75%
                root.update()
                time.sleep(0.03)

                # Aplicar los filtros
                df = platformFilter(df)
                progress_bar['value'] = 60  # 75%
                root.update()
                time.sleep(0.03)

                df = cookBesi(df)
                progress_bar['value'] = 100  # 100%
                root.update()

                return df
            
            else:
                progress_bar['value'] = 100  # 100%
                root.update()
                messagebox.showwarning("Advertencia", "El archivo no contiene el formato correcto")
                return None

        except Exception as e:
            progress_bar['value'] = 100  # 100%
            root.update()
            messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
        finally:
            time.sleep(0.03)
            progress_bar.destroy()
            root.update()
    else:
        progress_bar.stop()
        progress_bar.destroy()
        root.update()

#--------------------------------------------------------------
def uploadBom(root, contenedor_botones):
    # Encabezados esperados
    expected_headers = [
        'Plataforma',
        'Surtidor',
        'Turnos prod',
        'No. Part e SAS',	
        'No Parte Besi',
        'Descripcion',	
        'Capacidad en estanteria',
        'Std pack',
        'Inv',	
        'Estacion',
        'Estanteria',
        'Ubicación almacen',
        'Distancia'
    ]

    # Seleccionar archivo
    archivo = filedialog.askopenfilename(title='Subir BOM',filetypes=[("Archivos Excel", "*.xls;*.xlsx"), ("Todos los archivos", "*.*")])    

    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)
    
    progress_bar['value'] = 0  # Inicializa la barra a 0%
    root.update()

    if archivo:
        try:
            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 10  # 25%
            root.update()
            time.sleep(0.03)
            
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(archivo, header=1)

            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 20  # 25%
            root.update()
            time.sleep(0.03)

            # Asegurar que los encabezados sean strings y eliminar espacios en blanco
            df.columns = df.columns.map(lambda x: str(x).strip())
            # Limpiar encabezados (reemplazar saltos de línea y quitar espacios)
            df.columns = df.columns.str.replace('\n', ' ')  # Reemplazar saltos de línea con espacio
            df.columns = df.columns.str.strip()  # Eliminar espacios al inicio y final


            if set(expected_headers).issubset(df.columns):

                progress_bar['value'] = 40  # 75%
                root.update()
                time.sleep(0.03)

                # # Aplicar los filtros
                # df = platformFilter(df)
                # progress_bar['value'] = 60  # 75%
                # root.update()
                # time.sleep(0.03)

                progress_bar['value'] = 100  # 100%
                root.update()

                return df
            
            else:
                progress_bar['value'] = 100  # 100%
                root.update()
                messagebox.showwarning("Advertencia", "El archivo no contiene el formato correcto")
                return None

        except Exception as e:
            progress_bar['value'] = 100  # 100%
            root.update()
            messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
        finally:
            time.sleep(0.03)
            progress_bar.destroy()
            root.update()
    else:
        progress_bar.stop()
        progress_bar.destroy()
        root.update()

#--------------------------------------------------------------
def uploadLx02(root, contenedor_botones):

    # Encabezados esperados
    expected_headers = [
        'Material',
        'Centro',
        'Almacén',
        'Texto breve de material',
        'Tipo almacén',
        'Ubicación',
        'Stock disponible',
        'Unidad medida base',	
        'Fecha EM',
        'Unidad almacén',
        'Stock salida almacén',
        'Stock a entrar'
    ]

    warehouseTypes = [
        14,
        12,
        901,
        902,
        917,
        922,
        921,
    ]

    # Seleccionar archivo
    archivo = filedialog.askopenfilename(title='Subir LX02',filetypes=[("Archivos Excel", "*.xls;*.xlsx"), ("Todos los archivos", "*.*")])    

    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)
    
    progress_bar['value'] = 0  # Inicializa la barra a 0%
    root.update()

    if archivo:
        try:
            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 10  # 25%
            root.update()
            time.sleep(0.03)
            
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(archivo, header=0)

            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 20  # 25%
            root.update()
            time.sleep(0.03)

            # Asegurar que los encabezados sean strings y eliminar espacios en blanco
            df.columns = df.columns.map(lambda x: str(x).strip())
            # Limpiar encabezados (reemplazar saltos de línea y quitar espacios)
            df.columns = df.columns.str.replace('\n', ' ')  # Reemplazar saltos de línea con espacio
            df.columns = df.columns.str.strip()  # Eliminar espacios al inicio y final


            if set(expected_headers).issubset(df.columns):

                progress_bar['value'] = 40  # 75%
                root.update()
                time.sleep(0.03)
                df = df[df['Tipo almacén'].isin(warehouseTypes)]

                progress_bar['value'] = 100  # 100%
                root.update()

                return df
            
            else:
                progress_bar['value'] = 100  # 100%
                root.update()
                messagebox.showwarning("Advertencia", "El archivo no contiene el formato correcto")
                return None

        except Exception as e:
            progress_bar['value'] = 100  # 100%
            root.update()
            messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
        finally:
            time.sleep(0.03)
            progress_bar.destroy()
            root.update()
    else:
        progress_bar.stop()
        progress_bar.destroy()
        root.update()

#--------------------------------------------------------------
def uploadMData(root, contenedor_botones):

    # Encabezados esperados
    expected_headers = [
        'NP SAS',
        'NP VW',
        'Description',
        'Tipo de almacenamiento',
        'Supplier',
        'Planner',
        'Origin',
        'Politica  VWM',
        'Costo usd'
    ]

    # Seleccionar archivo
    archivo = filedialog.askopenfilename(title='Subir Mastar Data',filetypes=[("Archivos Excel", "*.xls;*.xlsx"), ("Todos los archivos", "*.*")])    

    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)
    
    progress_bar['value'] = 0  # Inicializa la barra a 0%
    root.update()

    if archivo:
        try:
            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 10  # 25%
            root.update()
            time.sleep(0.03)
            
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(archivo, header=0)

            # Aumentar el valor de la barra de progreso
            progress_bar['value'] = 20  # 25%
            root.update()
            time.sleep(0.03)


            # Asegurar que los encabezados sean strings y eliminar espacios en blanco
            df.columns = df.columns.map(lambda x: str(x).strip())

            # Limpiar encabezados (reemplazar saltos de línea y quitar espacios)
            df.columns = df.columns.str.replace('\n', ' ')  # Reemplazar saltos de línea con espacio
            df.columns = df.columns.str.strip()  # Eliminar espacios al inicio y final

            if set(expected_headers).issubset(df.columns):

                progress_bar['value'] = 40  # 75%
                root.update()
                time.sleep(0.03)

                progress_bar['value'] = 100  # 100%
                root.update()

                return df
            
            else:
                finish_destroy_progress(root, progress_bar)
                messagebox.showwarning("Advertencia", "El archivo no contiene el formato correcto")
                return None

        except Exception as e:
            progress_bar['value'] = 100  # 100%
            root.update()
            time.sleep(0.03 )
            messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
        finally:
            progress_bar.destroy()
            root.update()
    else:
        progress_bar.stop()
        progress_bar.destroy()
        root.update()

#-------------------------------------------------------------------
def strip_whitespace(x):
    if isinstance(x, str):
        return x.strip()
    else:
        return x

#-------------------------------------------------------------------
def calculateReport(root, contenedor_botones, besiDf,bomDf):

    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)
    
    # Se cálcula el paso para el loader en la iteración
    bomtaLen = len(bomDf)
    loadStep = 100 / bomtaLen
    
    try:
        reportHeaders = [
            'linea',
            'No Parte VW',
            'No Parte SAS',
            'Descripcion',	
            'Surtidor',	
            'Estacion',	
            'Estanteria',	
            'Ubicación almacen',	
            'Std pack',	
            'Inv',	
            'Capacidad cajas',	
            'Req diario Besi',	
            'Turnos',	
            'Req turno',	
            'Req hora',	
            'Cobertura x caja (hrs)',	
            'Cobertura x caja (min)',	
            'Cajas a surtir x turno',	
            'Distancia',	
            'Tiempo Surtimiento (Segundos) x caja',
            'Tiempo recorrido (Segundos) x caja',	
            'Work content x turno (min)',
            'Cajas x turno',
            'Cajas x hora',
            'Parcial 1',
            'Parcial 2',
            'Parcial 3',
        ]
        
        reportDf = pd.DataFrame(columns=reportHeaders)
        
        filas = []
        
        for i, row in bomDf.iterrows():

            #Asignar valores desde BOM
            sasNumberPart = row['No. Part e SAS']
            vwNumberPart = sasNumberPart.replace(' ','')

            #Filtrar conforme la referenciay obtener Dr
            platform = platforms[row['Plataforma']]
            reference = f"{platform['code']}-{vwNumberPart}"
            filterReference = besiDf[besiDf['Referencia'] == reference].copy()
            filterReference = filterReference.iloc[0]  if not filterReference.empty else None
            dr = filterReference['Dr'] if filterReference is not None else 0

            #Cálcular requerimiento por turno
            turns = row['Turnos prod']
            turnRequirement = math.ceil(dr / turns)

            #Cálcular requerimiento por hora
            hourRequeriment = math.ceil(turnRequirement / 8)

            #Definir standar pack
            stdPack = row['Std pack']

            #Cálcular Cobertura x caja (hrs)
            if not pd.isna(stdPack) and not pd.isna(hourRequeriment) and hourRequeriment != 0:
                boxHourCoberture = math.ceil(stdPack / hourRequeriment)
            else:
                boxHourCoberture = 0

            #Cálcular cobertura por caj aminutos
            if not pd.isna(stdPack) and not pd.isna(hourRequeriment) and hourRequeriment != 0:
                boxMinuteCobeture = math.ceil(stdPack / hourRequeriment * 60)
            else:
                boxMinuteCobeture = 0
            
            
            #Definir Inventario
            inv = row['Inv']

            #Cálcular capacidad de cajas
            if not pd.isna(stdPack) and not pd.isna(inv) and stdPack != 0:
                boxCapacity = inv / stdPack
            else:
                boxCapacity = 0

            #Cálcular Cajas a surtir x turno
            if not pd.isna(stdPack) and not pd.isna(turnRequirement) and turnRequirement != 0:
                boxTurnCoberture = math.ceil( turnRequirement/stdPack)
            else:
                boxTurnCoberture = 0

            #Cálcular cajas por hora
            boxPerHour = math.ceil(boxTurnCoberture / 8)

            #Definir distancia
            distance = row['Distancia']

            #Definir tiempo de surtimiento
            dispatchTime = 46
            
            #Cálcular tiempo recorrido (Segundos) * caja
            secondsPerBox = math.ceil(distance * 0.9 * 2) if not pd.isna(distance) else 0

            #Cálcular Work content x turno (min)
            workContent = round((dispatchTime + secondsPerBox) * boxTurnCoberture / 60, 2)

            new_row = {
                'linea': platform['label'],
                'No Parte VW': vwNumberPart,
                'No Parte SAS': sasNumberPart,
                'Descripcion': row['Descripcion'],
                'Surtidor': row['Surtidor'],
                'Estacion': row['Estacion'],
                'Estanteria': row['Estanteria'],
                'Ubicación almacen': row['Ubicación almacen'],
                'Std pack': stdPack,
                'Inv': inv,
                'Capacidad cajas': boxCapacity,
                'Req diario Besi': dr,
                'Turnos': turns,
                'Req turno': turnRequirement,
                'Req hora': hourRequeriment,
                'Cobertura x caja (hrs)': boxHourCoberture,
                'Cobertura x caja (min)': boxMinuteCobeture,
                'Cajas a surtir x turno': boxTurnCoberture,
                'Distancia': distance,
                'Tiempo Surtimiento (Segundos) x caja': dispatchTime,
                'Tiempo recorrido (Segundos) x caja': secondsPerBox,
                'Work content x turno (min)': workContent,
                'Cajas x turno': boxTurnCoberture,
                'Cajas x hora': boxPerHour,
                'Parcial 1': '',
                'Parcial 2': '',
                'Parcial 3': '',
            }
            filas.append(new_row)

            progress_bar['value'] = progress_bar['value'] + loadStep
            root.update()
            time.sleep(0.001)

        reportDf = pd.concat([reportDf, pd.DataFrame(filas)], ignore_index=True)
        
        reportDf = reportDf[reportDf['Cajas x turno'] > 0]
        
        return reportDf
    except Exception as e:
        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
        messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
    finally:
        progress_bar.destroy()
        root.update()

#----------------------------------------------------------------------------------------------
def openFile(fileName):
    
    time.sleep(1)  # Esperar 1 segundo

    # Intentar abrir el archivo
    try:
        os.startfile(fileName)
    except FileNotFoundError:
        messagebox.showwarning("Archivo no encontrado",f"Error: El archivo {fileName} no se pudo encontrar al intentar abrirlo.")
    return 


#-----------------------------------------------------------------
def exportReport(root, contenedor_botones,reportDf, REPORT_NAME = REPORT_NAME):

    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)

    if reportDf is None:
        messagebox.showwarning("Advertencia", "No se ha generado un reporte")
        finish_destroy_progress(root, progress_bar)
        return
    
    progress_bar['value'] = 10  # 100%
    root.update()
    time.sleep(0.03)
    
    # Pedir al usuario que seleccione la carpeta donde guardar el archivo
    folder_selected = filedialog.askdirectory(title="Selecciona la carpeta para guardar el archivo")

    progress_bar['value'] = 50  # 100%
    root.update()
    time.sleep(0.03)
    
    if not folder_selected:  # Si no se seleccionó ninguna carpeta
        print("No se seleccionó ninguna carpeta. Exportación cancelada.")
        finish_destroy_progress(root, progress_bar)
        return

    # Definir el nombre del archivo
    DOC_NAME = os.path.join(folder_selected, REPORT_NAME)
    try:
        # Guardar el DataFrame en un archivo Excel
        reportDf.to_excel(DOC_NAME, index=False)
        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
    except FileNotFoundError as e:
        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
        messagebox.showerror('Error al guardar archivo', f"No se pudo guardar el archivo: {e}")
        return
    except Exception as e:
        print(e)
        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
        messagebox.showerror('Error inesperado', f"Ocurrió un error al guardar el archivo: {e}")
        return
    finally:
        progress_bar.destroy()
        root.update()
        time.sleep(0.03)

    # Verificar si el archivo se creó
    if not os.path.exists(DOC_NAME):
        messagebox. print(f"Error: No se pudo crear el archivo {DOC_NAME}")
        return  # O manejar el error de la forma que desees

    print(f"Archivo guardado en: {os.path.abspath(DOC_NAME)}")  # Imprimir ruta absoluta

    openFile(DOC_NAME)

#-----------------------------------------------------------------
def exportReportMultiSh(root, contenedor_botones,reportLDfList, MULTI_REPORT_NAME = REPORT_NAME):


    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)

    if len(reportLDfList) == 0:
        messagebox.showwarning("Advertencia", "No se ha generado un reporte")
        finish_destroy_progress(root, progress_bar)
        return
    
    progress_bar['value'] = 10  # 100%
    root.update()
    time.sleep(0.03)
    
    # Pedir al usuario que seleccione la carpeta donde guardar el archivo
    folder_selected = filedialog.askdirectory(title="Selecciona la carpeta para guardar el archivo")

    progress_bar['value'] = 50  # 100%
    root.update()
    time.sleep(0.03)
    
    if not folder_selected:  # Si no se seleccionó ninguna carpeta
        print("No se seleccionó ninguna carpeta. Exportación cancelada.")
        finish_destroy_progress(root, progress_bar)
        return

    # Definir el nombre del archivo
    DOC_NAME = os.path.join(folder_selected, MULTI_REPORT_NAME)
    try:

        with pd.ExcelWriter(DOC_NAME) as writer:
            for df_info in reportLDfList:
                # Extraer el DataFrame y el nombre de la hoja
                df = df_info['df']
                sheet_name = df_info['sheetName']
                
                # Verificar que 'sheet_name' sea un string
                if not isinstance(sheet_name, str):
                    print(f"Error: El nombre de la hoja debe ser un string. Valor recibido: {sheet_name}")
                    continue  # Pasar al siguiente si el nombre de la hoja no es válido
                
                # Escribir el DataFrame en la hoja correspondiente
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
    except FileNotFoundError as e:
        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
        messagebox.showerror('Error al guardar archivo', f"No se pudo guardar el archivo: {e}")
        return
    except Exception as e:
        print(e)
        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
        messagebox.showerror('Error inesperado', f"Ocurrió un error al guardar el archivo: {e}")
        return
    finally:
        progress_bar.destroy()
        root.update()
        time.sleep(0.03)

    # Verificar si el archivo se creó
    if not os.path.exists(DOC_NAME):
        messagebox. print(f"Error: No se pudo crear el archivo {DOC_NAME}")
        return  # O manejar el error de la forma que desees

    print(f"Archivo guardado en: {os.path.abspath(DOC_NAME)}")  # Imprimir ruta absoluta

    openFile(DOC_NAME)

#-----------------------------------------------------------------------------------------
def cookDfToPdf(df):
    # Ordenar jerárquicamente por 'linea', 'Surtidor', y 'Cajas x turno'
    df = df.sort_values(by=['linea', 'Surtidor', 'Cajas x turno'], ascending=[True, True, False])

    cutReports = []

    # Iterar por cada valor único en 'linea'
    for linea_key in df['linea'].unique():
        lineDf = df[df['linea'] == linea_key]

        # Iterar por cada valor único en 'Surtidor'
        for surtidor_key in lineDf['Surtidor'].unique():
            
            surtidorDf = lineDf[lineDf['Surtidor'] == surtidor_key]

            # Añadir el DataFrame procesado a cutReports
            cutReports.append(surtidorDf)

    return cutReports

#--------------------------------------------------------------------------------------------
def exportPdfReport(root, contenedor_botones, reportDf):

    #crear la barra de carga
    progress_bar = createProgressBar(root, contenedor_botones)

    if reportDf is None:
        messagebox.showwarning("Advertencia", "No se ha generado un reporte")
        return
    
    progress_bar['value'] = 10
    root.update()
    time.sleep(0.03)
    
    try:
        # Cocinar los datos con la función previa
        cookedReport = cookDfToPdf(reportDf) 

        progress_bar['value'] = 50
        root.update()
        time.sleep(0.03)
        
        
        # Seleccionar carpeta de destino
        folder_selected = filedialog.askdirectory(title="Selecciona la carpeta para guardar el archivo")
        
        progress_bar['value'] = 60
        root.update()
        time.sleep(0.03)
        
        if not folder_selected:
            print("No se seleccionó ninguna carpeta. Exportación cancelada.")
            finish_destroy_progress(root, progress_bar)
            return

        REPORT_NAME = os.path.join(folder_selected, "ReporteSurtimiento.pdf")
        
        # Obtener la fecha actual
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        # Configurar el documento PDF
        pdf = SimpleDocTemplate(REPORT_NAME, pagesize=landscape(letter), leftMargin=5, rightMargin=5, topMargin=10, bottomMargin=10)
        elements = []
        
        # Obtener el tamaño de la página
        ancho, alto = landscape(letter)
        
        # Estilo de título
        styles = getSampleStyleSheet()

        # Añadir imagen y encabezado antes de cada DataFrame
        for df in cookedReport:

            if df.empty:
                continue  # Saltar DataFrames vacíos

            # Añadir imagen (opcional)
            if os.path.exists(IMAGE_PATH):
                img = Image(IMAGE_PATH, width=298/3, height=94/3)
                elements.append(img)
            else:
                messagebox.showwarning("Advertencia", f"La imagen no se encontró en {IMAGE_PATH}")

            # Añadir un título o encabezado
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"Reporte de surtimiento - {df['linea'].iloc[0]} - {df['Surtidor'].iloc[0]}", styles['Title']))  # Título dinámico basado en 'linea'
            
            df = df.drop([
                'linea',
                'No Parte SAS',
                'Surtidor',
                'Inv',
                'Req diario Besi',
                'Turnos',
                'Cobertura x caja (hrs)',
                'Cajas a surtir x turno',
                'Distancia',
                'Tiempo Surtimiento (Segundos) x caja',
                'Tiempo recorrido (Segundos) x caja',
                'Work content x turno (min)'
            ], axis=1)
            
            # # Reemplazar 'nan' por cadenas vacías y espacios por saltos de línea en las celdas del DataFrame
            df = df.map(lambda x: '' if pd.isna(x) else x)

            # Reemplazar espacios por saltos de línea en los nombres de las columnas
            df.columns = [col.replace(' ', '\n') for col in df.columns]
            
            # Convertir el DataFrame a una tabla
            data = [df.columns.tolist()] + df.values.tolist()

            table = Table(data)

            # Establecer el estilo de la tabla
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 3),
                ('TOPPADDING', (0, 0), (-1, 0), 3),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
                ('TOPPADDING', (0, 1), (-1, -1), 2),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ])
            table.setStyle(style)

            # Añadir la tabla al documento
            elements.append(Spacer(1, 12))
            elements.append(table)

            # Añadir un salto de página para que el siguiente DataFrame comience en la próxima página
            elements.append(PageBreak())

            # progress_bar['value'] = progress_bar['value'] + loadStep
            # root.update()
            # time.sleep(0.001)

        # Función para agregar el pie de página con la fecha
        def add_footer(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 8)
            canvas.drawString(inch, 0.12 * inch, f"Fecha: {current_date}")
            canvas.restoreState()

        # Guardar el archivo PDF
        pdf.build(elements,onLaterPages=add_footer, onFirstPage=add_footer)
    except PermissionError as e:
        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
        messagebox.showerror('Error de Permiso', f"El archivo '{REPORT_NAME}' está abierto o en uso. Ciérralo e intenta nuevamente.")
        return
    
    progress_bar['value'] = 100
    root.update()
    time.sleep(0.03)
    
    # Verificar si el archivo se creó
    if not os.path.exists(REPORT_NAME):
        messagebox.showerror(f"Error: No se pudo crear el archivo {REPORT_NAME}")
        return
    
    progress_bar.destroy()
    root.update()

    openFile(REPORT_NAME)

#-------------------------------------------------------------------
def calculateReport2(root, contenedor_botones, besiDf, lx02Df, mDataDf):
    
    
    # Extraer los encabezadois con formato de fecha
    dateHeaders = getDateHeaders(besiDf)

    # Definición de encabezados estáticos
    staticHeaders = [
        "NP SAS", "NP VW", "Description", "Supplier", "Planner", 
        "Origin", "Stock On Hand In Plant", "Days On Hand In Plant", 
        "DOH", "Politica  VWM"
    ]
    # Se guarda el número de encabezados
    headersLen = len(staticHeaders)

    # Se cálcula el paso para el loader en la iteración
    mDataLen = len(mDataDf)
    loadStep = 90 / mDataLen

    #Cálcular los días inhábiles
    probeDf = besiDf[besiDf['Noparte'] == '5NM857003FFLG'].copy()
    disabledDays = []
    for header in dateHeaders:
        #Cálcular días inhábiles basado en no parte general 857003
        dayRequeriment = sum(probeDf[header].to_list())
        disabledDays.append(False if dayRequeriment == 0 else True)

    # Se crea un array eliminando los dós últimos encabezados estáticos
    cookedHeaders = staticHeaders[0:headersLen - 2]

    # Se insertan los encabezados para los requerimientos del besi por día
    for header in dateHeaders:
        cookedHeaders.append(f"BESI {header}")

    # Se insertan los encabezados para el restante de piezasa despues del día
    for header in dateHeaders:
        cookedHeaders.append(f"Before {header}")

    # Se insertan los encabezados para el DOH Díario
    for header in dateHeaders:
        cookedHeaders.append(f"DOH {header}")

    # Se agregan los dos útimos ecabezados estáticos
    cookedHeaders = cookedHeaders + staticHeaders[headersLen - 2: headersLen]

    # Se aseguras hacer una lista contable
    reportHeaders = list(dict.fromkeys(cookedHeaders))

    # Se crea el dataframe del reporte 
    report2Df = pd.DataFrame(columns=reportHeaders)
    
    # Se crea un array para las filas
    filas = []
    progress_bar = createProgressBar(root,contenedor_botones)

    progress_bar['value'] = 10  # Inicializa la barra a 0%
    root.update()

    time.sleep(0.03)

    try:
        # Se recorre el Master Data
        for _, row in mDataDf.iterrows():

            # Setear numeros de parte
            sasNumberPart = row['NP SAS']
            vwNumberPart = row['NP VW']

            # Asignar descripción
            description = row['Description']

            # Asignar Supplier
            supplier = row['Supplier']

            # Asignar Planner
            planner = row['Planner']

            # Asignar Origin
            origin = row['Origin']

            # Extraer las filas que correspondan al número de parte
            besiFiltered = besiDf[besiDf['Noparte'] == vwNumberPart].copy()
            lx02Filtered = lx02Df[lx02Df['Material'] == sasNumberPart].copy()

            # Asignar tipo de almacenamiento
            storageType = row['Tipo de almacenamiento']

            if storageType == 'rack':
                lx02Filtered = lx02Filtered[lx02Filtered['Tipo almacén'] != 921]

            # Cálcular el inventario existente
            stockOnHandInPlant = sum(lx02Filtered['Stock disponible'].to_list())
            
            # Asignar Politica de VWM
            vwPolicy = row['Politica  VWM']

            newRow = {
                "NP SAS":sasNumberPart,
                "NP VW":vwNumberPart,
                "Description":description,
                "Supplier":supplier,
                "Planner":planner,
                "Origin":origin,
                "Stock On Hand In Plant":stockOnHandInPlant,
                "Politica  VWM":vwPolicy,
            }

            doh = 0
            sumRequeriment = 0
            
            for _, dateHader in enumerate(dateHeaders):

                #Cálcula el requerimiendo del besi por número de parte
                besiHeader = f"BESI {dateHader}"
                dayRequeriment = sum(besiFiltered[dateHader].to_list()) 
                sumRequeriment += dayRequeriment
                newRow[besiHeader] = dayRequeriment
                
                #Cálcula el restante despues del día
                beforeHeader = f"Before {dateHader}"
                startDatStock = stockOnHandInPlant
                stockOnHandInPlant -= dayRequeriment
                newRow[beforeHeader] = stockOnHandInPlant
                
                #Cálcula el restante despues del día
                dohHeader = f"DOH {dateHader}"

                if dayRequeriment > 0 and dayRequeriment <= startDatStock:
                    doh = doh + 1 

                if startDatStock > 0 and dayRequeriment > startDatStock:
                    doh = doh + ( startDatStock / dayRequeriment)

                newRow[dohHeader] = doh
            
                newRow['DOH'] = round(doh, 2)
                newRow['Days On Hand In Plant'] = round(doh, 2)
            
            progress_bar['value'] = progress_bar['value'] + loadStep
            root.update()
            time.sleep(0.03)

            if sumRequeriment > 0:
                filas.append(newRow)

        report2Df = pd.concat([report2Df, pd.DataFrame(filas)], ignore_index=True)
        
        progress_bar['value'] = 100  
        root.update()
        time.sleep(0.001)
        
        return report2Df
    except Exception as e:
        progress_bar['value'] = 100  # 100%
        root.update()
        messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
    finally:
        progress_bar.destroy()
        root.update()

#-----------------------------------------------------------------------------------------
def calculateCritics(dohDf, besiDf):
    critic_headers = [
        'NP SAS', 'Description', 'Supplier', 'Planner', 'Origin',	
        'Stock On Hand In Plant', 'Days On Hand In Plant', 'Politica  VWM'
    ]

    # Filtrar las columnas deseadas y hacer una copia explícita
    criticDf = dohDf[critic_headers].copy()
    
    # Asignar los origines esperados
    test1Origins = ['ASIA', 'EUROPA']
    test2Origins = ['FRONTERA', 'LOCAL', 'NACIONAL', 'USA']

    # Crear las condiciones
    test1 = (criticDf['Origin'].str.upper().isin(test1Origins)) & (criticDf['Days On Hand In Plant'] < 7)
    test2 = (criticDf['Origin'].str.upper().isin(test2Origins)) & (criticDf['Days On Hand In Plant'] < 1)

    # Asignar las condiciones al DataFrame
    criticDf.loc[:, 'test1'] = test1
    criticDf.loc[:, 'test2'] = test2

    # Test final
    criticDf.loc[:, 'test3'] = criticDf['test1'] | criticDf['test2']

    # Eliminar filas no críticas
    criticDf = criticDf[criticDf['test3']].copy()

    # Eliminar las columnas de prueba
    criticDf = criticDf.drop(['test1', 'test2', 'test3'], axis=1)

    # Agregar concatenación de plataformas
    criticDf['Plataformas'] = ''


    platformsDict = {
        '17A' : 'Jetta',
        '55N' : 'Tiguan',
        '2GM' : 'Taos',
        '57N' : 'Tayron'
    }

    for index, row in criticDf.iterrows():

        sasPartNumber = row['NP SAS']
        vwPartNumber = sasPartNumber.replace('\xa0', '').strip().replace(' ','').strip()

        print({sasPartNumber,vwPartNumber})

        filteredBesi = besiDf[besiDf['Noparte'] == vwPartNumber].copy()

        platforms = filteredBesi[['TME']].copy()
        platforms['Platform'] = platforms['TME'].apply(lambda x: platformsDict.get(x, 'Unknown')).astype(str)

        # Convertir la lista de plataformas a una cadena separada por guiones
        platform_string = '-'.join(platforms['Platform'].to_list())
        
        # Asignar la cadena al DataFrame `criticDf` en la columna 'Plataformas'
        criticDf.at[index, 'Plataformas'] = platform_string

    return criticDf

#-----------------------------------------------------------------------------------------
def calculateExcesses(dohDf, mDataDf):
    critic_headers = [
        'NP SAS', 'Description', 'Supplier', 'Planner', 'Origin',	
        'Stock On Hand In Plant', 'DOH promedio', 'Politica  VWM'
    ]

    # Filtrar las columnas que contienen 'DOH' en el nombre
    doh_headers = dohDf.columns[dohDf.columns.str.contains('BESI ', case=False)]

    # Filtrar las columnas deseadas y hacer una copia explícita
    excessDf = dohDf.copy()

    # Calcular el promedio de las columnas que contienen 'DOH' en su nombre, excluyendo ceros
    excessDf['DOH promedio'] = excessDf['Stock On Hand In Plant'] / excessDf[doh_headers].replace(0, pd.NA).mean(axis=1)

    # Filtramos las columnas deseadas 
    excessDf = excessDf[critic_headers]

    # Insertar la columna en una posición específica
    excessDf.insert(6, 'DOH promedio', excessDf.pop('DOH promedio'))

    # Asignar los origines esperados
    test1Origins = ['ASIA', 'EUROPA']
    test2Origins = ['FRONTERA', 'LOCAL', 'NACIONAL', 'USA']

    # Declartación de test
    test1 = (excessDf['Origin'].str.upper().isin(test1Origins)) & (excessDf['DOH promedio'] > excessDf['Politica  VWM'])
    test2 = (excessDf['Origin'].str.upper().isin(test2Origins)) & (excessDf['DOH promedio'] > excessDf['Politica  VWM'])

    # Agregamos los test al df
    excessDf.loc[:,'test1'] = test1 
    excessDf.loc[:,'test2'] = test2 

    # Test final
    excessDf.loc[:,'test3'] = excessDf['test1'] | excessDf['test2']

    # Eliminar las filas que no son excedentes
    excessDf = excessDf[excessDf['test3']]

    # Limpiar columnas
    excessDf = excessDf.drop(['test1','test2', 'test3'], axis = 1)

    # Seleccionar solo las columnas relevantes de `mDataDf`
    mDataCosto = mDataDf[['NP SAS', 'Costo usd']].copy()

    # Realizar el merge en la columna `NP SAS` para añadir `Costo usd` a `excessDf`
    excessDf = excessDf.merge(mDataCosto, on='NP SAS', how='left')

    # Renombrar la columna a `Costo unitario` si deseas mantener ese nombre
    excessDf = excessDf.rename(columns={'Costo usd': 'Costo unitario'})
    
    # Cálculamos 
    excessDf['Valor exceso'] =  (excessDf['Stock On Hand In Plant'] / excessDf['DOH promedio']) * ( excessDf['DOH promedio'] - excessDf['Politica  VWM']) * excessDf['Costo unitario'] 
    
    # Cálcular total del costo
    totalValue = sum(excessDf['Valor exceso'].to_list())

    # Agregamos el Valor total
    lastRow = {'Costo unitario': 'Valor Total', 'Valor exceso' : totalValue}
    excessDf = pd.concat([excessDf, pd.DataFrame([lastRow])], ignore_index = True)

    return excessDf

#-------------------------------------------------------------------
def calculateLx02Report1(lx02Df):

    lx02Copy = lx02Df.copy()

    uniquesMaterials = lx02Copy['Material'].unique()

    lx02Report1Headers = ['NP SAS', 'Cuenta de Ubicación'] 
    lx02Report1 = pd.DataFrame(columns=lx02Report1Headers)

    rows = []

    for material in uniquesMaterials:

        materialDf = lx02Copy[lx02Copy['Material'] == material]
        materialCount = len(materialDf)

        if materialCount <= 2 :
            newRow = {'NP SAS': material, 'Cuenta de Ubicación' : materialCount}
            rows.append(newRow)

    lx02Report1 = pd.concat([lx02Report1,pd.DataFrame(rows)], ignore_index=True)

    return lx02Report1

#------------------------------------------------------------------
def calculateLx02Report2(lx02Df, mDataDf):
    # Especifica las columnas clave de cada DataFrame
    columna_lx02 = 'Material'  # Columna en lx02Df para comparar
    columna_mData = 'NP SAS'  # Columna en mDataDf para comparar

    # Filtrar las filas de `lx02Df` que no tienen valores en `columna_mData` de `mDataDf`
    resultado = lx02Df[~lx02Df[columna_lx02].isin(mDataDf[columna_mData])]

    # Tomar solo las columnas de interes
    finalCols = ['Material', 'Stock disponible']

    resultado = resultado[finalCols]
    
    # Eliminar duplicados para obtener una fila por cada valor único de Material
    resultado = resultado.drop_duplicates(subset=['Material'])
    return resultado

#-------------------------------------------------------------------
def calculateReport3(root, contenedor_botones, dohDf, lx02Df, mDataDf, besiDf):
    
    progress_bar = createProgressBar(root,contenedor_botones)

    progress_bar['value'] = 10  # Inicializa la barra a 0%
    root.update()
    time.sleep(0.03)

    dataframeList = []
    try:

        # Cálculo del reporte de criticos
        criticsDf = calculateCritics(dohDf, besiDf)
        critic_report = {'df':criticsDf, 'sheetName' : 'Reporte de cortos'}

        # Cálculo del reporte de excesos
        excessesDf = calculateExcesses(dohDf, mDataDf)
        excess_report = {'df' : excessesDf,'sheetName' : 'Reporte de excesos'}

        # Filtrar lx02 con tomando solo almacen 014
        filteredLx02Df = lx02Df[lx02Df['Tipo almacén'] == 14].copy()

        # Cálculo del report Componentes con max 2 bines de existencia
        lx02Report1Df = calculateLx02Report1(filteredLx02Df)
        lx02_report_1 = {'df' : lx02Report1Df,'sheetName' : 'Comp max 2 bines de existencia'}

        # Cálculo del report Componentes con max 2 bines de existencia
        lx02Report2Df = calculateLx02Report2(filteredLx02Df, mDataDf)
        lx02_report_2 = {'df' : lx02Report2Df,'sheetName' : 'Comp no incluidos en MD'}
        
        dataframeList.append(critic_report)
        dataframeList.append(excess_report)
        dataframeList.append(lx02_report_1)
        dataframeList.append(lx02_report_2)

        progress_bar['value'] = 100  # 100%
        root.update()
        time.sleep(0.03)
    except Exception as e:
        progress_bar['value'] = 100  # 100%
        root.update()
        print(e)
        messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
    finally:
        progress_bar.destroy()
        root.update()
    
    exportReportMultiSh(root,contenedor_botones,dataframeList,REPORT_NAME_3)