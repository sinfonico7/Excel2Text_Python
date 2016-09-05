from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from openpyxl import load_workbook
from openpyxl import Workbook
from tkinter import messagebox


# metodo que ejecuta el buscador de archivos
def buscar_archivo():
    global ruta_archivo
    buscar_archivo = askopenfilename(filetypes=(("Archivos Excel", "*.xlsx"),("All files", "*.*")))
    ruta_archivo = buscar_archivo
    ruta_origen.insert(0, ruta_archivo)



    return 0


# metodo que setea el destino donde ira el archivo
def setear_destino():
    global destino
    buscar_destino = askdirectory()
    destino = buscar_destino
    ruta_destino.insert(0, destino)
    return 0


# metodo que creara un archivo de texto formateado como libro diario
def conversion():

    ruta_archivo = ruta_origen.get()
    print(ruta_archivo)
    if ruta_archivo != "":

        libro = load_workbook(ruta_archivo)
        hojas =libro.get_sheet_names()
        llenarComboBox(hojas)
        messagebox.showinfo("Felicitaciones!", "Logré Abrir archivo de excel!")

    else:
        messagebox.showinfo("Advertencia", "No ha seleccionado Ningun Archivo")
    return 0

# llenamos el combo box con los datos del libro tomado por el entry de origen
def llenarComboBox(paginas):
    combo_paginas.config(values=paginas)
    return 0


maestro = Tk()
maestro.geometry()

# sector de eleccion de archivos
frame_archivos = LabelFrame(maestro, text="primer frame", pady=42)

# widgets de frame eleccion de archivos

texto_origen = Label(frame_archivos, text="Sleeccione Archivo")
texto_destino = Label(frame_archivos, text="Sleeccione Destino")
ruta_origen = Entry(frame_archivos, width=50)
ruta_destino = Entry(frame_archivos, width=50)

# este boton quisiera que invocara el metodo busqueda de archivos, pero no se como invocar el comando
boton_origen = Button(frame_archivos, text="origen", command=buscar_archivo)  # command= buscar_archivo()
boton_destino = Button(frame_archivos, text="destino", command=setear_destino)

# distribucion de lugar de los wigets
frame_archivos.grid(row=0, column=0)
texto_origen.grid(row=0, column=0)
ruta_origen.grid(row=1, column=0)
boton_origen.grid(row=1, column=1)
texto_destino.grid(row=3, column=0)
ruta_destino.grid(row=4, column=0)
boton_destino.grid(row=4, column=1)

# ----------------------------------------------------------------------------------
# sector de tipo de libro
frame_tipoLibro = LabelFrame(maestro, text="Tipo libro", width=300, height=110, padx=60, pady=5)





# widgets
combo_libros = ttk.Combobox(frame_tipoLibro)
combo_meses = ttk.Combobox(frame_tipoLibro)
combo_anios = ttk.Combobox(frame_tipoLibro)
combo_paginas = ttk.Combobox(frame_tipoLibro)

texto_tipo_libro = Label(frame_tipoLibro, text="Seleccione Tipo de Libro")
texto_mes_salida = Label(frame_tipoLibro, text="Seleccione un Mes")
texto_anio_salida = Label(frame_tipoLibro, text="Seleccione un Anio")
texto_nombre_pagina = Label(frame_tipoLibro, text="Seleccione una Pagina")

# distribucion de los widgets
frame_tipoLibro.grid(row=0, column=1)
texto_tipo_libro.pack()
combo_libros.pack()
texto_mes_salida.pack()
combo_meses.pack()
texto_anio_salida.pack()
combo_anios.pack()
texto_nombre_pagina.pack()
combo_paginas.pack()
# valores de combos
combo_libros.config(values=('Libro_Diario', 'Libro_Mayor', 'Libro_Balance'))
combo_meses.config(values=(
    'Enero', 'Feberero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre',
    'Diciembre'))
combo_anios.config(values=('2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020'))

# ----------------------------------------------------------------------------------
# sector de imagenes
frame_imagenes = LabelFrame(maestro, text="Imagenes", width=50, height=50)
# widgets


# distribucion de los widgets
frame_imagenes.grid(row=1, column=0)

# ----------------------------------------------------------------------------------
# sector de comenzar
frame_comienzo_proceso = LabelFrame(maestro, text="Comienzo Proceso")
# widgets
boton_comenzar = Button(frame_comienzo_proceso, text="Comenzar", command=conversion)
# distribucion de los widgets
frame_comienzo_proceso.grid(row=1, column=1)
boton_comenzar.pack()

maestro.mainloop()



