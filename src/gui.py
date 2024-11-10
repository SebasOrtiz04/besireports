import tkinter as tk
from tkinter import ttk
import pandas as pd


from data.paths import ICON_PATH, REPORT_NAME
from data.index import platforms
from utils.main import uploadBesi, uploadBom, calculateReport, exportReport, exportPdfReport

# Etiquetas para mostrar el número de filas en cada pestaña
row_count_label_besi = None
row_count_label_bom = None
row_count_label_report = None

#Dataframes de archivos
besiDf = None
bomDf = None
reportDf = None

#-------------------------------------------------------------------
def createReport(report_treeview, notebook):

    global reportDf

    if besiDf is None:
        return

    if bomDf is None:
        return
    
    reportDf = calculateReport(besiDf,bomDf)
        
    if reportDf is None:
        return
        
    notebook.select(2)
    
    for i in report_treeview.get_children():
        report_treeview.delete(i)

    report_treeview["columns"] = list(reportDf.columns)
    report_treeview["show"] = "headings"

    for col in reportDf.columns:
        report_treeview.heading(col, text=col)

    for _, row in reportDf.iterrows():
        report_treeview.insert("", "end", values=list(row))

    row_count_label_report.config(text=f"Número de filas REPORTE: {len(reportDf)}")

    exportPdfReport(reportDf)

    

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
    
    createReport(report_treeview, notebook)

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
    
    createReport(report_treeview, notebook)

#----------------------------------------------------------------------------
def createGui():
    global row_count_label_besi, row_count_label_bom, row_count_label_report
    
    root = tk.Tk()
    root.title("Análissis BESI")
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
    report_book = ttk.Frame(notebook)

    # Añadir las pestañas al notebook
    notebook.add(besi_book, text="BESI")
    notebook.add(bom_book, text="BOM")
    notebook.add(report_book, text="REPORTE")

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

    # Crear el Treeview para el REPORTE
    report_treeview_frame = ttk.Frame(report_book)
    report_treeview_frame.pack(fill="both", expand=True)
    report_treeview = ttk.Treeview(report_treeview_frame)
    scrollbar_report = ttk.Scrollbar(report_treeview_frame, orient="horizontal", command=report_treeview.xview)
    report_treeview.configure(xscrollcommand=scrollbar_report.set)
    report_treeview.pack(side="top", fill="both", expand=True)
    scrollbar_report.pack(side="bottom", fill="x")

    # Crear etiquetas para mostrar el número de filas en cada pestaña
    row_count_label_besi = tk.Label(besi_book, text="Número de filas BESI: 0", font=("Arial", 10))
    row_count_label_besi.pack(pady=5)

    row_count_label_bom = tk.Label(bom_book, text="Número de filas BOM: 0", font=("Arial", 10))
    row_count_label_bom.pack(pady=5)

    row_count_label_report = tk.Label(report_book, text="Número de filas REPORTE: 0", font=("Arial", 10))
    row_count_label_report.pack(pady=5)

    # Añadir Botón para cargar besi
    boton_cargar_besi = ttk.Button(contenedor_botones, text="SUBIR BESI", command=lambda: besiToDf(root, contenedor_botones, besi_treeview, notebook, report_treeview))
    boton_cargar_besi.pack(side="left", padx=10)

    # Añadir botón para cargar BOM
    boton_cargar_bom = ttk.Button(contenedor_botones, text="SUBIR BOM", command=lambda: bomToDf(root, contenedor_botones, bom_treeview, notebook, report_treeview))
    boton_cargar_bom.pack(side="left", padx=10)

    #Añadir botón para descargar excel
    boton_cargar_bom = ttk.Button(contenedor_botones, text="DESCARGAR .xlsx", command=lambda: exportReport(reportDf))
    boton_cargar_bom.pack(side="left", padx=10)

    #Añadir botón para descargar pdf
    boton_cargar_bom = ttk.Button(contenedor_botones, text="DESCARGAR .pdf", command=lambda: exportPdfReport(reportDf))
    boton_cargar_bom.pack(side="left", padx=10)

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
