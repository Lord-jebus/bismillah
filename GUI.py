# -*- coding: utf-8 -*-
"""
Created on Sat Apr 12 23:22:16 2025

@author: Jebus
"""
import sys
import os
import tkinter as tk
from tkinter import Tk,Checkbutton,IntVar
from tkinter import ttk
from tkinter import filedialog
import filtrado

filename = ""
tipo = True
file_source = ""
filtro = ""

def get_base_path():
    if getattr(sys, 'frozen', False):
        # Si es un .exe creado por PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Si es un script .py
        return os.path.dirname(os.path.abspath(__file__))

def recurso_relativo(ruta_relativa):
    """Retorna la ruta válida para usar recursos, ya sea en desarrollo o en el ejecutable."""
    if hasattr(sys, '_MEIPASS'):
        # Si está empaquetado con PyInstaller
        base_path = sys._MEIPASS
    else:
        # Si se está ejecutando desde el script directamente
        base_path = os.path.abspath(".")

    return os.path.join(base_path, ruta_relativa)

class Checkbox(ttk.Checkbutton):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.variable = tk.BooleanVar(self)
        self.configure(variable=self.variable)
    
    def checked(self):
        return self.variable.get()
    
    def check(self):
        self.variable.set(True)
    
    def uncheck(self):
        self.variable.set(False)

class Aplicacion(ttk.Frame):
    global filename, file_source, filtro
    def __init__(self, parent):
        super().__init__(parent)
        
        self.img_boton = tk.PhotoImage(file=recurso_relativo("images/acerca2.png"))
        self.boton_acercade = ttk.Button(
            text="Información", image=self.img_boton, compound=tk.TOP,
            command=self.abrir_ventana_secundaria)
        self.boton_acercade.place(x=400, y=20)
        
        self.img_boton3 = tk.PhotoImage(file=recurso_relativo("images/buscar4.png"))
        self.boton_convertir = ttk.Button(
            text="Archivo", image=self.img_boton3, compound=tk.TOP,
            command=self.abrir_archivo)
        self.boton_convertir.place(x=20, y=20)

        self.img_boton2 = tk.PhotoImage(file=recurso_relativo("images/procesar3.png"))
        self.boton_procesar = ttk.Button(
            text="Generar Resumen", image=self.img_boton2, compound=tk.TOP,
            command=self.procesar)
        self.boton_procesar.place(x=370, y=305)
        
        self.etiqueta_file_name = ttk.Label(
            parent, font=("Tahoma", 12, "bold"),
            text="Archivo de referencia para filtrar:", foreground="#0070C0")
        self.etiqueta_file_name.place(x=20, y=120)
        
        self.etiqueta_filename = ttk.Label(
            parent, font=("Tahoma", 8, "bold"), foreground="#0070C0",
            text="Ruta: ")
        self.etiqueta_filename.place(x=20, y=145)

        self.etiqueta_fuente = ttk.Label(
            parent, font=("Tahoma", 12, "bold"), foreground="#0070C0",
            text="Fuente: ")
        self.etiqueta_fuente.place(x=20, y=170)

        self.etiqueta_criterios = ttk.Label(
            parent, font=("Tahoma", 12, "bold"),
            text="Criterios de búsqueda", foreground="#0070C0",)
        self.etiqueta_criterios.place(x=20, y=260)

        self.etiqueta_keywords = ttk.Label(
            parent, font=("Tahoma", 12, "bold"), foreground="#0070C0",
            text="Palabra clave de interés:")
        self.etiqueta_keywords.place(x=20, y=200)
        
        self.entry_var = tk.StringVar()        
        self.entry = ttk.Entry(textvariable=self.entry_var, width=40)
        self.entry.place(x=230, y=200)

        self.val = IntVar()
        self.checkbox = tk.Checkbutton(self,
            text="Author Keywords / Keywords Author",
                font=("Tahoma", 12),  
                fg="blue", variable=self.val)
        self.checkbox.place(x=42, y=300)        
        self.place(width=400, height=400)
        
        self.val2 = IntVar()
        self.checkbox2 = tk.Checkbutton(self,
            text="Indexed Keywords / Keywords Plus",
                font=("Tahoma", 12),  
                fg="blue", variable=self.val2)
        self.checkbox2.place(x=42, y=330)
        self.place(width=400, height=400)

        self.val3 = IntVar()
        self.checkbox3 = tk.Checkbutton(self,
            text="Abstract",
                font=("Tahoma", 12),  
                fg="blue", variable=self.val3)
        self.checkbox3.place(x=42, y=360)
        self.place(width=400, height=400) 

        self.etiqueta_advertencia = ttk.Label(
            parent, font=("Tahoma", 12, "bold"),
            text="")
        self.etiqueta_advertencia.place(x=20, y=230)

        self.etiqueta_archivo = ttk.Label(
            parent, font=("Tahoma", 10), foreground="blue",
            text="")
        self.etiqueta_archivo.place(x=20, y=430)

        self.etiqueta_ruta = ttk.Label(
            parent, font=("Tahoma", 8), foreground="blue",
            text="")
        self.etiqueta_ruta.place(x=20, y=400)

    def abrir_ventana_secundaria(self):
        self.ventana_secundaria = VentanaSecundaria()
        
    def abrir_archivo(self):
        global file_source, filename
        #seleccionar archivo
        filename = filedialog.askopenfilename()
        self.etiqueta_filename.config(
            text=f"Ruta: {filename}")
        if filename[-4:] == '.txt':
            file_source = "WOS"
            self.etiqueta_fuente.config(text="Fuente: Web of Science (Clarivate)")
        elif filename[-4:] == '.csv':
            self.etiqueta_fuente.config(text="Fuente: Scopus (ELSEVIER)")
            file_source = "SCO"
        else:
            self.etiqueta_fuente.config(text="Archivo no compatible para su procesamiento", foreground="#660033")
            
    def procesar(self):
        global file_source, filtro
        filtro = self.entry_var.get()

        # Obtener el estado de los checkbox
        v1 = self.val.get()
        v2 = self.val2.get()
        v3 = self.val3.get()

        # Validación de combinaciones inválidas
        if not v1 and not v2 and not v3:
            self.etiqueta_advertencia.config(
                text="Seleccione un criterio de filtro", foreground="#660033")
            return

        if v3 and (v1 or v2):
            self.etiqueta_advertencia.config(
                text="No puede combinar el tercer criterio con los otros dos", foreground="#660033")
            return

        # Determinación del modo válido
        if v1 and v2:
            modo = 'OR'
        elif v1:
            modo = '1'
        elif v2:
            modo = '2'
        elif v3:
            modo = '3'
        else:
            # Fallback por seguridad, aunque no debería llegar aquí
            self.etiqueta_advertencia.config(
                text="Selección inválida de filtros", foreground="#660033")
            return

        # Selección de función de procesamiento
        procesar_func = {
            "WOS": filtrado.procesar_wos,
            "SCO": filtrado.procesar_sco
        }.get(file_source)

        if procesar_func:
            fecha = procesar_func(filename, filtro, modo=modo)
            print(fecha)
            self.etiqueta_ruta.config(
                text=recurso_relativo("images")[:-7], foreground="#660033")
            self.etiqueta_archivo.config(text=fecha, foreground="#660033")
        else:
            self.etiqueta_advertencia.config(
                text="Archivo inválido", foreground="#660033")

        # Actualiza la etiqueta de ruta después del procesamiento
        self.etiqueta_ruta.config(text=get_base_path(), foreground="#660033")
      
        
              
    def procesar1(self):
        global file_source
        print(file_source)
        if self.val.get() == False and self.val2.get() == False:            
            self.etiqueta_advertencia.config(text="Seleccione un criterio de filtro", foreground="#660033")            
        else:
            if self.val.get() and self.val2.get():
                if file_source == "WOS":
                    filtrado.procesar_wos(filename, filtro, modo = 'AND')
                if  file_source == "SCO":
                    filtrado.procesar_sco(filename, filtro, modo = 'AND')
                else:
                    self.etiqueta_advertencia.config(text="Archivo inválido", foreground="#660033")
            else:
                if file_source == "WOS":
                    filtrado.procesar_wos(filename, filtro, modo = 'OR')
                if  file_source == "SCO":
                    filtrado.procesar_sco(filename, filtro, modo = 'OR')
                else:
                    self.etiqueta_advertencia.config(text="Archivo inválido", foreground="#660033")
                
class SCOPUS1(ttk.Frame):    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=800, height=450)

        self.label = ttk.Label(
            self, font=("Tahoma", 12),
            text="Si exporta un archivo de CSV desde SCOPUS\nseleccione las opciones 'Author keywords' e 'Indexed keywords'", foreground="#0070C0")
        self.label.pack()

        self.sc1 = tk.PhotoImage(file=recurso_relativo("images/SCOPUS1.png"))
        self.etiqueta3 = tk.Label(self, image=self.sc1)
        self.etiqueta3.pack()#place(x=1, y=1)
        
        self.sc2 = tk.PhotoImage(file=recurso_relativo("images/SCOPUS2.png"))
        self.etiqueta4 = tk.Label(self, image=self.sc2)        
        self.etiqueta4.pack()#place(x=100, y=100)
        
class WOS1(ttk.Frame):    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=800, height=450)
        self.label = ttk.Label(
            self, font=("Tahoma", 12),
            text='Si exporta un archivo de texto sin formato desde Web of Science\nseleccione las opciones "KeyWords" y "KeyWords Plus"', foreground="#0070C0")
        self.label.pack()
        
        # Generar el contenido de cada una de las pestañas.
        self.ws1 = tk.PhotoImage(file=recurso_relativo("images/WOS1.png"))
        self.etiqueta = tk.Label(self, image=self.ws1)
        self.etiqueta.pack()#place(x=1, y=1)
        
        self.ws2 = tk.PhotoImage(file=recurso_relativo("images/WOS2.png"))
        self.etiqueta2 = tk.Label(self, image=self.ws2)
        self.etiqueta2.pack()#place(x=100, y=100)


class SCOPUS(ttk.Frame):    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=800, height=450)

        # Crear canvas y scrollbar
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Crear un frame dentro del canvas
        self.content_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        # Actualizar región del canvas según el contenido
        self.content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # ---- CONTENIDO ----
        label = ttk.Label(
            self.content_frame, font=("Tahoma", 12),
            text="Si exporta un archivo de CSV desde SCOPUS\nseleccione las opciones 'Author keywords' e 'Indexed keywords'",
            foreground="#0070C0"
        )
        label.pack(pady=10)

        self.sc1 = tk.PhotoImage(file=recurso_relativo("images/SCOPUS1.png"))
        etiqueta3 = tk.Label(self.content_frame, image=self.sc1)
        etiqueta3.pack(pady=10)

        self.sc2 = tk.PhotoImage(file=recurso_relativo("images/SCOPUS2.png"))
        etiqueta4 = tk.Label(self.content_frame, image=self.sc2)
        etiqueta4.pack(pady=10)


class WOS(ttk.Frame):    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.config(width=800, height=450)

        # Crear el canvas con scrollbar
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Posicionar canvas y scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Crear el frame interno donde irá todo el contenido
        self.content_frame = ttk.Frame(canvas)
        
        # Asociar el frame al canvas usando una window
        canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        # Actualizar la región del scroll cada vez que se cambie el tamaño del contenido
        self.content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # ---- CONTENIDO (como el que tú tenías) ----
        label = ttk.Label(
            self.content_frame, font=("Tahoma", 12),
            text='Si exporta un archivo de texto sin formato desde Web of Science\nseleccione las opciones "KeyWords" y "KeyWords Plus"',
            foreground="#0070C0"
        )
        label.pack(pady=10)

        self.ws1 = tk.PhotoImage(file=recurso_relativo("images/WOS1.png"))
        etiqueta1 = tk.Label(self.content_frame, image=self.ws1)
        etiqueta1.pack(pady=10)

        self.ws2 = tk.PhotoImage(file=recurso_relativo("images/WOS2.png"))
        etiqueta2 = tk.Label(self.content_frame, image=self.ws2)
        etiqueta2.pack(pady=10)

class AboutFrame(ttk.Frame):
    global version, fecha
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=800, height=450)
        
        content = ttk.Frame(self)
        content.place(relx=0.5, rely=0.5, anchor="center")
        
        label = ttk.Label(
            content, font=("Tahoma", 12, "bold"),
            text="Diseño conceptual:", foreground="#1070F0")
        label.pack(pady=5)
        #self.label.pack(expand=True, fill="both")

        label1 = ttk.Label(
            content, font=("Arial", 12, "bold"),
            text="Juan Pablo Franco R", foreground="#0070C0")
        label1.pack(pady=5)
        #self.label.pack(expand=True, fill="both")

        label2 = ttk.Label(
            content, font=("Tahoma", 12, "bold"),
            text="Desarrollado por:", foreground="#1070F0")
        label2.pack(pady=5)
        #self.label.pack(expand=True, fill="both")

        label3 = ttk.Label(
            content, font=("Arial", 12, "bold"),
            text="Jesús Martínez V", foreground="#0070C0")
        label3.pack(pady=5)
        #self.label.pack(expand=True, fill="both")

        label4 = ttk.Label(
            content, font=("Arial", 10),
            text=f"Version {version}", foreground="#0070C0")
        #self.label.pack(expand=True, fill="both")
        label4.pack(pady=5)

        label5 = ttk.Label(
            content, font=("Arial", 10),
            text=f"Fecha: {fecha}", foreground="#0070C0")
        label5.pack(pady=5)
        #self.label.pack(expand=True, fill="both")
        
   
class VentanaSecundaria(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=800, height=500)#720*480
        self.title("Acerca de:")
        self.resizable(False, False)
        icono = tk.PhotoImage(file=recurso_relativo("images/ventana2.png"))
        self.iconphoto(False, icono)

        # Generar el panel de pestañas.
        self.notebook = ttk.Notebook(self)

        self.greeting_frame = SCOPUS(self.notebook)
        self.notebook.add(
            self.greeting_frame, text="SCOPUS", padding=10)
        
        self.about_frame = WOS(self.notebook)
        self.notebook.add(
            self.about_frame, text="Web of Science", padding=10)

        self.about_frame = AboutFrame(self.notebook)
        self.notebook.add(
            self.about_frame, text="Acerca de", padding=10)
        
        self.notebook.pack(fill="both", expand=True)

        self.focus()
        self.grab_set()

version = "1.1"
fecha = "Abril 2.025"
ventana = tk.Tk()
ventana.title(f"Bismillah - Filtro WOS / Scopus")
ventana.config(width=500, height=460)
ventana.resizable(False, False)
icono = tk.PhotoImage(file=recurso_relativo("images/ventana2.png"))
ventana.iconphoto(False, icono)
# Puedes probar: 'default', 'alt', 'classic', 'vista', 'xpnative', 'clam'
style = ttk.Style(ventana)
style.theme_use("vista")

app = Aplicacion(ventana)
ventana.mainloop()
