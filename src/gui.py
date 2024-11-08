import tkinter as tk
from tkinter import ttk
import pandas as pd


from data.paths import ICON_PATH, REPORT_NAME_2
from data.index import platforms
from utils.main import uploadBesi, uploadBom, calculateReport, exportPdfReport, uploadLx02, calculateReport2, uploadMData, exportReport, calculateReport3

# Etiquetas para mostrar el número de filas en cada pestaña
row_count_label_besi = None
row_count_label_bom = None
row_count_label_lx02 = None
row_count_label_mData = None
row_count_label_report = None
row_count_label_report2 = None

#Dataframes de archivos
besiDf = None
bomDf = None
lx02Df = None
mDataDf = None
reportDf = None
report2Df = None

#-------------------------------------------------------------------
def createReport2(root, report2_treeview, notebook, contenedor_botones):

    global report2Df, besiDf, reportDf, mDataDf, lx02Df

    if besiDf is None or lx02Df is None or reportDf is None or mDataDf is None:
        return

    
    report2Df = calculateReport2(root, contenedor_botones, besiDf,lx02Df, mDataDf)
        
    if report2Df is None:
        return
        
    notebook.select(5)
    
    for i in report2_treeview.get_children():
        report2_treeview.delete(i)

    report2_treeview["columns"] = list(report2Df.columns)
    report2_treeview["show"] = "headings"

    for col in report2Df.columns:
        report2_treeview.heading(col, text=col)

    for _, row in report2Df.iterrows():
        report2_treeview.insert("", "end", values=list(row))

    row_count_label_report2.config(text=f"Número de filas REPORTE: {len(report2Df)}")

    calculateReport3(root, contenedor_botones,report2Df,lx02Df, mDataDf)

#-------------------------------------------------------------------
def createReport(root, contenedor_botones,report_treeview, notebook):

    global reportDf

    if besiDf is None:
        return

    if bomDf is None:
        return
    
    reportDf = calculateReport(root, contenedor_botones,besiDf,bomDf)
        
    if reportDf is None:
        return
        
    notebook.select(4)
    
    for i in report_treeview.get_children():
        report_treeview.delete(i)

    report_treeview["columns"] = list(reportDf.columns)
    report_treeview["show"] = "headings"

    for col in reportDf.columns:
        report_treeview.heading(col, text=col)

    for _, row in reportDf.iterrows():
        report_treeview.insert("", "end", values=list(row))

    row_count_label_report.config(text=f"Número de filas REPORTE: {len(reportDf)}")

    exportPdfReport(root, contenedor_botones,reportDf)

#-------------------------------------------------------------------
def besiToDf(root, contenedor_botones, besi_treeview, notebook, report_treeview):
    notebook.select(0)
    
    global row_count_label_besi, besiDf

    besiDf = uploadBesi(root, contenedor_botones)

    if besiDf is not None:
        for i in besi_treeview.get_children():
            besi_treeview.delete(i)

        besi_treeview["columns"] = list(besiDf.columns)
        besi_treeview["show"] = "headings"

        for col in besiDf.columns:
            besi_treeview.heading(col, text=col)

        for _, row in besiDf.iterrows():
            besi_treeview.insert("", "end", values=list(row))

        row_count_label_besi.config(text=f"Número de filas BESI: {len(besiDf)}")
    
    createReport(root, contenedor_botones,report_treeview, notebook)

#------------------------------------------------------------------------
def bomToDf(root, contenedor_botones, bom_treeview, notebook, report_treeview):
    notebook.select(1)
    
    global row_count_label_bom, bomDf
    
    bomDf = uploadBom(root, contenedor_botones)

    if bomDf is not None:
        for i in bom_treeview.get_children():
            bom_treeview.delete(i)

        bom_treeview["columns"] = list(bomDf.columns)
        bom_treeview["show"] = "headings"

        for col in bomDf.columns:
            bom_treeview.heading(col, text=col)

        for _, row in bomDf.iterrows():
            bom_treeview.insert("", "end", values=list(row))

        row_count_label_bom.config(text=f"Número de filas BOM: {len(bomDf)}")
    
    createReport(root, contenedor_botones,report_treeview, notebook)

#------------------------------------------------------------------------
def lx02ToDf(root, contenedor_botones, lx02_treeview, notebook):
    notebook.select(2)
    
    global row_count_label_lx02, lx02Df
    
    lx02Df = uploadLx02(root, contenedor_botones)

    if lx02Df is not None:
        for i in lx02_treeview.get_children():
            lx02_treeview.delete(i)

        lx02_treeview["columns"] = list(lx02Df.columns)
        lx02_treeview["show"] = "headings"

        for col in lx02Df.columns:
            lx02_treeview.heading(col, text=col)

        for _, row in lx02Df.iterrows():
            lx02_treeview.insert("", "end", values=list(row))

        row_count_label_lx02.config(text=f"Número de filas LX02: {len(lx02Df)}")

#------------------------------------------------------------------------
def mDataToDf(root, contenedor_botones, mData_treeview, notebook, report2_treeview):
    notebook.select(3)
    
    global row_count_label_mData, mDataDf
    
    mDataDf = uploadMData(root, contenedor_botones)

    if mDataDf is not None:
        for i in mData_treeview.get_children():
            mData_treeview.delete(i)

        mData_treeview["columns"] = list(mDataDf.columns)
        mData_treeview["show"] = "headings"

        for col in mDataDf.columns:
            mData_treeview.heading(col, text=col)

        for _, row in mDataDf.iterrows():
            mData_treeview.insert("", "end", values=list(row))

        row_count_label_mData.config(text=f"Número de filas Master Data: {len(mDataDf)}")
    
    createReport2(root, report2_treeview, notebook, contenedor_botones)

#----------------------------------------------------------------------------
def createGui():
    global row_count_label_besi, row_count_label_bom, row_count_label_report, row_count_label_lx02, row_count_label_report2, row_count_label_mData
    
    root = tk.Tk()
    root.title("Motherson Surtido de cajas diario")
    root.iconbitmap(ICON_PATH)
    root.geometry("800x600")

    # Configurar el estilo
    style = ttk.Style(root)
    style.theme_use('winnative')

    # Estilizar botones
    style.configure('TButton', font=('Arial', 8,'bold'), padding=3)
    style.map('TButton',
              foreground=[('pressed', 'white'), ('active', 'gray')],
              background=[('pressed', 'white'), ('active', 'gray')])

    # Cambiar el padding de las pestañas
    style.configure('TNotebook.Tab', padding=[10, 2], font=('Arial', 10, 'bold'))

    # Crear el widget Notebook
    notebook = ttk.Notebook(root)
    notebook.grid(row=2, column=0, padx=10, pady=1, sticky="nsew")

    # Crear el contenedor de pestañas
    besi_book = ttk.Frame(notebook)
    bom_book = ttk.Frame(notebook)
    lx02_book = ttk.Frame(notebook)
    mData_book = ttk.Frame(notebook)
    report_book = ttk.Frame(notebook)
    report2_book = ttk.Frame(notebook)

    # Añadir las pestañas al notebook
    notebook.add(besi_book, text="BESI")
    notebook.add(bom_book, text="BOM")
    notebook.add(lx02_book, text="LX02")
    notebook.add(mData_book, text="MD")
    notebook.add(report_book, text="Surtido cajas")
    notebook.add(report2_book, text="DOH")

    # Crear el contenedor para los botones
    contenedor_botones = tk.Frame(root)
    contenedor_botones.grid(row=0, column=0, padx=20, pady=10, sticky="w")

    # Crear el contenedor para los botones
    contenedor_etiquetas = tk.Frame(root)
    contenedor_etiquetas.grid(row=1, column=0, padx=20, pady=10, sticky="w")

    # Crear el Treeview para BESI
    besi_treeview_frame = ttk.Frame(besi_book)
    besi_treeview_frame.pack(fill="both", expand=True)
    besi_treeview = ttk.Treeview(besi_treeview_frame)
    scrollbar_besi = ttk.Scrollbar(besi_treeview_frame, orient="horizontal", command=besi_treeview.xview)
    besi_treeview.configure(xscrollcommand=scrollbar_besi.set)
    besi_treeview.pack(side="top", fill="both", expand=True)
    scrollbar_besi.pack(side="bottom", fill="x")

    # Crear el Treeview para BOM
    bom_treeview_frame = ttk.Frame(bom_book)
    bom_treeview_frame.pack(fill="both", expand=True)
    bom_treeview = ttk.Treeview(bom_treeview_frame)
    scrollbar_bom = ttk.Scrollbar(bom_treeview_frame, orient="horizontal", command=bom_treeview.xview)
    bom_treeview.configure(xscrollcommand=scrollbar_bom.set)
    bom_treeview.pack(side="top", fill="both", expand=True)
    scrollbar_bom.pack(side="bottom", fill="x")

    # Crear el Treeview para LX02
    lx02_treeview_frame = ttk.Frame(lx02_book)
    lx02_treeview_frame.pack(fill="both", expand=True)
    lx02_treeview = ttk.Treeview(lx02_treeview_frame)
    scrollbar_bom = ttk.Scrollbar(lx02_treeview_frame, orient="horizontal", command=lx02_treeview.xview)
    lx02_treeview.configure(xscrollcommand=scrollbar_bom.set)
    lx02_treeview.pack(side="top", fill="both", expand=True)
    scrollbar_bom.pack(side="bottom", fill="x")

    # Crear el Treeview para Master Data
    mData_treeview_frame = ttk.Frame(mData_book)
    mData_treeview_frame.pack(fill="both", expand=True)
    mData_treeview = ttk.Treeview(mData_treeview_frame)
    scrollbar_bom = ttk.Scrollbar(mData_treeview_frame, orient="horizontal", command=mData_treeview.xview)
    mData_treeview.configure(xscrollcommand=scrollbar_bom.set)
    mData_treeview.pack(side="top", fill="both", expand=True)
    scrollbar_bom.pack(side="bottom", fill="x")


    # Crear el Treeview para el REPORTE
    report_treeview_frame = ttk.Frame(report_book)
    report_treeview_frame.pack(fill="both", expand=True)
    report_treeview = ttk.Treeview(report_treeview_frame)
    scrollbar_report = ttk.Scrollbar(report_treeview_frame, orient="horizontal", command=report_treeview.xview)
    report_treeview.configure(xscrollcommand=scrollbar_report.set)
    report_treeview.pack(side="top", fill="both", expand=True)
    scrollbar_report.pack(side="bottom", fill="x")

    # Crear el contenedor para los botones
    report_buttons = tk.Frame(report_treeview_frame)
    report_buttons.pack(side="top", pady=10)

    #Añadir botón para descargar excel
    boton_cargar_bom = ttk.Button(report_buttons, text="Surtido cajas.xlsx", command=lambda: exportReport(root, contenedor_botones,reportDf))
    boton_cargar_bom.pack(side="left", padx=10)

    #Añadir botón para descargar pdf
    boton_cargar_bom = ttk.Button(report_buttons, text="Surtido cajas.pdf", command=lambda: exportPdfReport(root, contenedor_botones,reportDf))
    boton_cargar_bom.pack(side="left", padx=10)

    # Crear el Treeview para el REPORTE 2
    report2_treeview_frame = ttk.Frame(report2_book)
    report2_treeview_frame.pack(fill="both", expand=True)
    report2_treeview = ttk.Treeview(report2_treeview_frame)
    scrollbar_report = ttk.Scrollbar(report2_treeview_frame, orient="horizontal", command=report2_treeview.xview)
    report2_treeview.configure(xscrollcommand=scrollbar_report.set)
    report2_treeview.pack(side="top", fill="both", expand=True)
    scrollbar_report.pack(side="bottom", fill="x")

    # Crear el contenedor para los botones
    report2_buttons = tk.Frame(report2_treeview_frame)
    report2_buttons.pack(side="top", pady=10)

    #Añadir botón para descargar excel
    boton_cargar_bom = ttk.Button(report2_buttons, text="DOH.xlsx", command=lambda: exportReport(root, contenedor_botones,report2Df,REPORT_NAME_2))
    boton_cargar_bom.pack(side="left", padx=10)

    #Añadir botón para descargar pdf
    boton_cargar_bom = ttk.Button(report2_buttons, text="Análisis DOH.xlsx", command=lambda: calculateReport3(root, contenedor_botones,report2Df, lx02Df, mDataDf))
    boton_cargar_bom.pack(side="left", padx=10)

    # Crear etiquetas para mostrar el número de filas en cada pestaña
    row_count_label_besi = tk.Label(besi_book, text="Número de filas BESI: 0", font=("Arial", 10))
    row_count_label_besi.pack(pady=5)

    row_count_label_bom = tk.Label(bom_book, text="Número de filas BOM: 0", font=("Arial", 10))
    row_count_label_bom.pack(pady=5)

    row_count_label_lx02 = tk.Label(lx02_book, text="Número de filas XL02: 0", font=("Arial", 10))
    row_count_label_lx02.pack(pady=5)

    row_count_label_mData = tk.Label(mData_book, text="Número de filas Master Data: 0", font=("Arial", 10))
    row_count_label_mData.pack(pady=5)

    row_count_label_report = tk.Label(report_book, text="Número de filas Surtido Cajas: 0", font=("Arial", 10))
    row_count_label_report.pack(pady=5)

    row_count_label_report2 = tk.Label(report2_book, text="Número de filas REPORTE: 0", font=("Arial", 10))
    row_count_label_report2.pack(pady=5)


#----Botones---------------------------------------------------------------------------------

    # Añadir Botón para cargar besi
    boton_cargar_besi = ttk.Button(contenedor_botones, text="SUBIR BESI", command=lambda: besiToDf(root, contenedor_botones, besi_treeview, notebook, report_treeview))
    boton_cargar_besi.pack(side="left", padx=10)

    # Añadir botón para cargar BOM
    boton_cargar_bom = ttk.Button(contenedor_botones, text="SUBIR BOM", command=lambda: bomToDf(root, contenedor_botones, bom_treeview, notebook, report_treeview))
    boton_cargar_bom.pack(side="left", padx=10)

    # Añadir botón para cargar LX02
    boton_cargar_lx02 = ttk.Button(contenedor_botones, text="SUBIR LX02", command=lambda: lx02ToDf(root, contenedor_botones, lx02_treeview, notebook))
    boton_cargar_lx02.pack(side="left", padx=10)

    # Añadir botón para cargar LX02
    boton_cargar_mData = ttk.Button(contenedor_botones, text="SUBIR MD", command=lambda: mDataToDf(root, contenedor_botones, mData_treeview, notebook, report2_treeview))
    boton_cargar_mData.pack(side="left", padx=10)

    label = tk.Label(contenedor_etiquetas, text='Plataformas:', font=("Arial", 10))
    label.pack(side="left", padx=10)

    for key in platforms.keys():
        platform_label = tk.Label(contenedor_etiquetas, text=platforms[key]['label'], font=("Arial", 10, "bold"))
        platform_label.pack(side="left", padx=10)

    # Configuración de filas y columnas
    root.grid_rowconfigure(0, weight=0)  # Sin expansión vertical para botones
    root.grid_rowconfigure(1, weight=0)  # Sin expansión vertical para etiquetas
    root.grid_rowconfigure(2, weight=1)  # Expansión vertical para el notebook

    # Configuración de columnas para que se expandan
    root.grid_columnconfigure(0, weight=1)  # Expansión horizontal para la columna

    root.mainloop()

#----------------------------------------------------
if __name__ == "__main__":
    createGui()
