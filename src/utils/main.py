import pandas as pd
import re
import time
import math
import os

from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from data.index import platforms
from data.paths import REPORT_NAME, PDF_REPORT_NAME, IMAGE_PATH,REPORT_NAME_2

from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image, PageBreak
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from tkinter import filedialog, messagebox
from datetime import datetime

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

#-------------------------------------------------------------------
def strip_whitespace(x):
    if isinstance(x, str):
        return x.strip()
    else:
        return x

#-------------------------------------------------------------------
def calculateReport(besiDf,bomDf):
    
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
        vwNumberPart = row['No. Part e SAS']
        sasNumberPart = vwNumberPart.replace(' ','')

        #Filtrar conforme la referenciay obtener Dr
        platform = platforms[row['Plataforma']]
        reference = f"{platform['code']}-{sasNumberPart}"
        filterReference = besiDf[besiDf['Referencia'] == reference]
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

    reportDf = pd.concat([reportDf, pd.DataFrame(filas)], ignore_index=True)
    
    reportDf = reportDf[reportDf['Cajas x turno'] > 0]
    
    return reportDf

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
def exportReport(reportDf, REPORT_NAME = REPORT_NAME):

    if reportDf is None:
        messagebox.showwarning("Advertencia", "No se ha generado un reporte")
        return
    
    # Pedir al usuario que seleccione la carpeta donde guardar el archivo
    folder_selected = filedialog.askdirectory(title="Selecciona la carpeta para guardar el archivo")

    if not folder_selected:  # Si no se seleccionó ninguna carpeta
        print("No se seleccionó ninguna carpeta. Exportación cancelada.")
        return

    # Definir el nombre del archivo
    DOC_NAME = os.path.join(folder_selected, REPORT_NAME)
    try:
        # Guardar el DataFrame en un archivo Excel
        reportDf.to_excel(DOC_NAME, index=False)
    except FileNotFoundError as e:
        messagebox.showerror('Error al guardar archivo', f"No se pudo guardar el archivo: {e}")
        return
    except Exception as e:
        print(e)
        messagebox.showerror('Error inesperado', f"Ocurrió un error al guardar el archivo: {e}")
        return

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
def exportPdfReport(reportDf):
    if reportDf is None:
        messagebox.showwarning("Advertencia", "No se ha generado un reporte")
        return

    # Cocinar los datos con la función previa
    cookedReport = cookDfToPdf(reportDf)

    # Seleccionar carpeta de destino
    folder_selected = filedialog.askdirectory(title="Selecciona la carpeta para guardar el archivo")

    if not folder_selected:
        print("No se seleccionó ninguna carpeta. Exportación cancelada.")
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
        df = df.applymap(lambda x: '' if pd.isna(x) else x)

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

    # Función para agregar el pie de página con la fecha
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.drawString(inch, 0.12 * inch, f"Fecha: {current_date}")
        canvas.restoreState()

    try:
        # Guardar el archivo PDF
        pdf.build(elements,onLaterPages=add_footer, onFirstPage=add_footer)
    except PermissionError as e:
        messagebox.showerror('Error de Permiso', f"El archivo '{REPORT_NAME}' está abierto o en uso. Ciérralo e intenta nuevamente.")
        return
    except Exception as e:
        messagebox.showerror('Error inesperado', f"Ocurrió un error al guardar el archivo: {e}")
        return

    # Verificar si el archivo se creó
    if not os.path.exists(REPORT_NAME):
        messagebox.showerror(f"Error: No se pudo crear el archivo {REPORT_NAME}")
        return

    openFile(REPORT_NAME)

#-------------------------------------------------------------------
def calculateReport2(root, contenedor_botones, besiDf, lx02Df, mDataDf):
    
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
    probeDf = besiDf[besiDf['Noparte'] == '5NM857003FFLG']
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
    
    progress_bar['value'] = 10  # Inicializa la barra a 0%
    root.update()
    time.sleep(0.03)

    try:
        # Se recorre el Master Data
        for i, row in mDataDf.iterrows():

            # Setear numeros de parte
            vwNumberPart = row['NP SAS']
            sasNumberPart = row['NP VW']

            # Asignar descripción
            description = row['Description']

            # Asignar Supplier
            supplier = row['Supplier']

            # Asignar Planner
            planner = row['Planner']

            # Asignar Origin
            origin = row['Origin']

            # Extraer las filas que correspondan al número de parte
            besiFiltered = besiDf[besiDf['Noparte'] == vwNumberPart]
            lx02Filtered = lx02Df[lx02Df['Material'] == sasNumberPart]

            # Asignar tipo de almacenamiento
            storageType = row['Tipo de almacenamiento']

            if storageType == 'rack':
                lx02Filtered = lx02Filtered[lx02Filtered['Tipo almacén'] == 921]
                print(lx02Filtered)

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

            
            for index, dateHader in enumerate(dateHeaders):

                #Cálcula el requerimiendo del besi por número de parte
                besiHeader = f"BESI {dateHader}"
                dayRequeriment = sum(besiFiltered[dateHader].to_list()) 
                newRow[besiHeader] = dayRequeriment
                
                #Cálcula el restante despues del día
                beforeHeader = f"Before {dateHader}"
                stockOnHandInPlant -= dayRequeriment
                newRow[beforeHeader] = stockOnHandInPlant
                
                #Cálcula el restante despues del día
                dohHeader = f"DOH {dateHader}"

                if stockOnHandInPlant > 0 and dayRequeriment > 0 and dayRequeriment <= stockOnHandInPlant:
                    doh = doh + 1 

                if stockOnHandInPlant > 0 and dayRequeriment > stockOnHandInPlant:
                    doh = doh + ( stockOnHandInPlant / dayRequeriment)

                newRow[dohHeader] = doh
            
                newRow['DOH'] = round(doh, 2)
                newRow['Days On Hand In Plant'] = round(doh, 2)
            
            progress_bar['value'] = progress_bar['value'] + loadStep
            root.update()
            time.sleep(0.03)

            filas.append(newRow)

        report2Df = pd.concat([report2Df, pd.DataFrame(filas)], ignore_index=True)
        
        progress_bar['value'] = 100  
        root.update()
        time.sleep(0.03)

        exportReport(report2Df, REPORT_NAME_2)
        
        return report2Df
    except Exception as e:
        progress_bar['value'] = 100  # 100%
        root.update()
        messagebox.showerror("Algo salió mal", f"Hubo un error al cargar el archivo: {str(e)}")
    finally:
        progress_bar.destroy()
        root.update()