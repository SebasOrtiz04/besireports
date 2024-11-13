
import pandas as pd
import time

from tkinter import filedialog, messagebox
from tkinter import filedialog, messagebox
from controllers.main_controller import createProgressBar

class ExcelService:

    @staticmethod
    #--------------------------------------------------------------
    def uploadExcel(root, contenedor_botones, expected_headers, title):

        # Seleccionar archivo
        archivo = filedialog.askopenfilename(title,filetypes=[("Archivos Excel", "*.xls;*.xlsx"), ("Todos los archivos", "*.*")])    

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