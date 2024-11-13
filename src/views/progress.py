import time

import tkinter as tk
from tkinter import ttk

class ProgressBar:
    def __init__(self, root, contenedor_botones):
        # Guardar referencias a root y contenedor
        self.root = root
        self.contenedor_botones = contenedor_botones
        
        # Configurar el estilo personalizado para la barra de progreso
        self.style = ttk.Style(self.root)
        self.style.configure("my.Horizontal.TProgressbar", 
                             troughcolor='white',  
                             background='red',    
                             thickness=10)
        
        # Crear la barra de progreso en modo determinate
        self.progress_bar = ttk.Progressbar(self.contenedor_botones, 
                                            style="my.Horizontal.TProgressbar", 
                                            mode='determinate', 
                                            length=200)
        self.progress_bar.pack(side="left", padx=10)
        self.progress_bar['value'] = 0  # Inicializa la barra a 0%
        self.root.update()

    def set_value(self, value):
        """Establece el valor de la barra de progreso."""
        self.progress_bar['value'] = value
        self.root.update()
        time.sleep(0.03)

    def get_progress_bar(self):
        """Devuelve la instancia de la barra de progreso si es necesario acceder a ella desde afuera."""
        return self.progress_bar
