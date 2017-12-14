#Librerias
import pandas as pd
import numpy as np
import threading
import time
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as mb

#Funciones
def hilo():
    '''
    Crea un hilo para la apertura del archivo
    '''
    hilo1 = threading.Thread(target=leer)
    hilo1.start()
	
def hilo2():
    '''
    Crea un hilo para crear los reportes
    '''
    hilo2 = threading.Thread(target=reporte)
    hilo2.start()
    
def cerrar():
    mb.showinfo("Cargando...", "La ventana se cerrara automáticamente cuando termine la carga del archivo.")
	
def leer():
    '''
    Abre el archivo excel seleccionado
    '''
    #Ruta del archivo excel que se leera
    ruta = filedialog.askopenfilename(title="Abrir",filetypes = (("Fichero Excel","*.xlsx"),("Fichero Excel 97-2003","*.xls")))
    if(ruta != ''):
        global df
        v1 = Toplevel(root)
        v1.title('Report v1.0')
        v1.resizable(0,0)
        v1.geometry('100x50')
        v1.grab_set()
        v1.protocol("WM_DELETE_WINDOW", cerrar) #para evitar el cierre de la ventana
        #Barra de progreso
        Label(v1, text='Cargando...').place(x=10, y=0)
        pb = ttk.Progressbar(v1, mode='indeterminate')
        pb.place(x=10,y=20)
        #Iniciar barra de progreso
        pb.start()
        #Leer archivo
        xlsx = pd.ExcelFile(ruta)
        df = pd.read_excel(xlsx, 'Actividades_Promotores_Detalle',usecols=[5,9,12,13,20,25,27])
        #Crear valores para el spinbox
        global territorio
        region = df['Nuevo territorio Regional'].dropna()
        territorio = list(set(list(region)))
        s1.config(values=territorio)
        #Detener barra de progreso
        pb.stop()
        b1.destroy()
        v1.destroy()
        s1.place(x=10,y=10)
        b2.place(x=10,y=50)
    else:
        mb.showwarning('Error','Ningún archivo seleccionado')

def reporte():
    '''
    Crear el reporte del territorio seleccionado en un archivo excel
    Una tabla por hoja
    '''
    v1 = Toplevel(root)
    v1.title('Report v1.0')
    v1.resizable(0,0)
    v1.geometry('100x50')
    v1.grab_set()
    v1.protocol("WM_DELETE_WINDOW", cerrar) #para evitar el cierre de la ventana
    #Barra de progreso
    Label(v1, text='Cargando...').place(x=10, y=0)
    pb = ttk.Progressbar(v1, mode='indeterminate')
    pb.place(x=10,y=20)
    #Iniciar barra de progreso
    pb.start()
    #Empiezan reportes    
    region = s1.get()
    df1 = df['Nuevo territorio Regional']== region
    df2 = df[df1]
    #Filtramos las columnas necesarias
    df1 = df2[['TIENDA','RESPUESTA','COMENTARIOS','Nombre Asesor','Obligatoria']]
    #Filtramos los asesores existentes y se gurdan en un conjunto para evitar valores repetidos
    asesores = set(df1['Nombre Asesor'])
    #Filtramos las respuestas existentes y se guardan en un conjunto para evitar valores repetidos
    respuestas = set(df1['RESPUESTA'])
    #Creamos un diccionario donde se gurdaran los valores de cada asesor por cada respuesta
    dict = {}
    #Creamos una lista para cada asesor para que se guarden sus datos correspondientes
    for col in asesores:
        dict[col] = []

    #Obtenemos los valores para cada respuesta de cada asesor
    for i in respuestas:
        asesor_respuesta = df1['RESPUESTA'] == i
        for j in asesores:
            asesor = df1['Nombre Asesor'] == j
            maritza = df1[asesor]
            #Guardamos los valores de cada asesor en su correspondiente lista
            dict[j].append(len(maritza[asesor_respuesta]))

    #Creamos el indice del DataFrame por cada respuesta
    indice = list(respuestas)
    #Creamos el dataframe a partir del diccionario y el indice creados anteriormente
    dataFrame = pd.DataFrame(dict, index=indice)
    #Sumamos las filas y agregamos una nueva columna
    ndf = dataFrame.copy()
    ndf[region] = dataFrame.sum(1)
    ndf
    #sumamos las columnas
    f = ndf.sum(0)
    #Agregamos las sumas de columna como una nueva fila
    final = ndf.append(f, ignore_index=True)
    #Como se eliminan los indices, los creamos otra vez agregando el indice 'Total' faltante para las sumas por columna
    lista = list(respuestas)
    lista.append('Total')
    #Agregamos el indice
    final.index = lista
    #Termina primer  tabla
    #Segunda tabla asesores, tienda, respuesta aceptable
    quitarRespuestas = (df1['RESPUESTA'] != 'SI, CON material (Banderin,banderola,poster)') & (df1['RESPUESTA'] != 'SI')
    quitar = df1[quitarRespuestas]
    t2 = quitar[['Nombre Asesor','TIENDA','TIENDA']]
    t2.columns = ['Nombre Asesor', 'TIENDA','# RESPUESTAS NO ACEPTABLES']
    grupo = t2.groupby(['Nombre Asesor','TIENDA'])['# RESPUESTAS NO ACEPTABLES'].count()
    #Regresar un grupo a un dataframe
    t = grupo.reset_index()
    tf = t.sort_values(by=['# RESPUESTAS NO ACEPTABLES'], ascending=False)
    tf.index = range(1,len(tf)+1)
    #Exportamos a un archivo en hojas diferentes
    fecha = time.strftime("%Y%m%d")
    nombre = str(fecha)+str(s1.get())+'.xlsx'
    archivo = pd.ExcelWriter(nombre)
    final.to_excel(archivo, sheet_name='tabla1')
    tf.to_excel(archivo, sheet_name='tabla2')
    archivo.save()
    #Detener barra de progreso
    pb.stop()
    v1.destroy()
	
#Ventana principal
root = Tk()

#Propiedades Ventanas
root.title('Report v1.0')
root.resizable(0,0)
root.geometry('200x200')

#Botones
territorio = None
df = None
	#Abrir archivo

b1 = Button(root,text="Abrir archivo", command=hilo, width='10')
b1.place(x=10,y=10)
	#Seleccionar territorio
l1 = Label(root, text="Cargando...")
#l1.place(x=10,y=60)
s1 = Spinbox(root, values=territorio)
#s1.place(x=10,y=80)
	#Crear reporte
b2 = Button(root, text='Reporte', command=hilo2)
#b2.place(x=10,y=120)
	

#Loop Principal
root.mainloop()

