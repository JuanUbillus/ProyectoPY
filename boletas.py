from ast import Pass, Str
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tokenize import Double
import PyPDF2,re,fitz,pyodbc
from PyPDF2 import PdfFileReader, PdfFileWriter,PdfFileMerger
from tkinter import filedialog
import os
from ctypes.wintypes import PINT
from pickle import APPEND
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io
import pandas as pd
from shutil import rmtree

#conexion a la Base de Datos y crear la  lista de ordenamiento
def conexion_bd():
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=DCGURKAS;DATABASE=GRUPO_GURKAS;UID=sa;PWD=Gurkas2019')
    cursor = cnxn.cursor()
    cod_unidad = txt_cod_unidad.get()
    cursor_migracion = cnxn.cursor()
    query =  " exec sp_order_boletas '" + str(cod_unidad) +"'"
    df = pd.read_sql(query,cnxn)
    cursor.close()
    del cursor
    cnxn.close()
    for x in df['DOCT_IDENT'].tolist():
        lista_de_documento_ordenado.append(x)
    print(lista_de_documento_ordenado)
    print("The total number of elements in the list: ", len(lista_de_documento_ordenado))
    return(lista_de_documento_ordenado)


#crear un diccionario para realizar los filtros de las boletas
def datos_cbo():   
    for key in lista_unidades:
        for value in lista_unidades_codigo:
            Undiad_Key[key] = value
            lista_unidades_codigo.remove(value)
            break

#almacenar la ubicaion del logotipo de la empresa 
def abrir_archivo_logo():
    ubicacion_file=filedialog.askopenfilename(initialdir = "/"
    , title ="Seleccione el logo de la empresa",filetypes = ((".jpg","*.*"),
    ("all files","*.*")))
    logo = (ubicacion_file)
    ruta_logo.append(logo)
    messagebox.showinfo(message="Se selecciono correctamente el logotipo de la empresa", title="Exitoso")

#almacenar la ubicacion de la firma de la empresa
def abrir_archivo_firma():
    ubicacion_file=filedialog.askopenfilename(initialdir = "/"
    , title ="Seleccione la firma",filetypes = ((".jpg","*.*"),
    ("all files","*.*")))
    firma= (ubicacion_file)
    ruta_firma.append(firma)
    messagebox.showinfo(message="Se selecciono correctamente la fimra de la empresa", title="Exitoso")

#abrir carpeta del archivo principal
def abrir_archivo():
    #ubicamos y sacamos la ruta del archivo 
    ubicacion_file=filedialog.askopenfilename(initialdir = "/"
    , title ="Seleccione archivo",filetypes = ((".pdf","*.*"),
    ("all files","*.*")))
    ruta = (ubicacion_file)
    folder = os.path.split(ruta)[0]
    file_os = os.path.split(ruta)[1]
    #sacamos el numero de paginas del pdf
    pdf = PdfFileReader(open(ruta,'rb'))
    numero_pages = pdf.getNumPages()
    print("file : ",ruta,"pages",numero_pages)
    if os.path.isdir(os.path.join(folder,"Boletas") +'_PLAME'):
        tk.messagebox.showwarning(title="error", message="la carpeta "+file_os+" ya existe")
    else :
        tk.messagebox.showinfo(title="correcto", message="se importo exitosamente la cantidad de "+str(numero_pages)+" PDF")
        ruta_folder.append(folder)
        ruta_pdf.append(folder)
        return folder, file_os

def armado_pdf_final():
    ruta_nombre_pdf = ruta_pdf[0]
    folder_cedula = os.path.split(ruta_nombre_pdf)[0]
    folder = os.path.split(ruta_nombre_pdf)[0]
    file_os = os.path.split(ruta_nombre_pdf)[1]
    #creamos la carpeta
    carpeta_destino = os.mkdir(os.path.join(ruta_nombre_pdf) +'/'+'Boletas_Ordenada') 
    carpeta = os.path.join(ruta_nombre_pdf)+'/'+'Boletas_Ordenada'
    carpeta_2 = os.path.join(ruta_nombre_pdf)+'/Boletas_PLAME'
    merger = PdfFileMerger()

    for doc in lista_de_documento_ordenado:        
        try:
            merger.append(PdfFileReader(carpeta_2+"/"+doc+".pdf"))
            merger.append(PdfFileReader(carpeta_2+"/"+doc+".pdf"))
            print("Presonal Titular de la unidad : "+doc)
        except:
            print("Personal no encontrado en la Unidad : "+ doc)
    merger.write(carpeta+"/"+"Boletas_"+ "Boletas_Ordenado"+".pdf")
    messagebox.showinfo(message="Proceso de Boletas Terminado", title="Exitoso")

#Escribir en el pdf el neto 
def insertar_neto_pdf():
    for item in lista_pdf_individual:
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        valor = neto_pagado[item.split("/")[4]]
        can.drawString(312,60, "TOTAL NETO A PAGAR DEL MES : "+ str(valor))
        can.save()
        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        pdf_ubicacion = os.path.join(str(lista_pdf_nuevo_unido_parte_1[0]))
        pdf = PdfFileReader(pdf_ubicacion)
        existing_pdf = PdfFileReader(item, "rb")
        pdfWriter = PdfFileWriter()
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        pdfWriter.addPage(page)
        outputStream = open(item, "wb")
        pdfWriter.write(outputStream)
        outputStream.close()

        pdf_ubicacion = os.path.join(item)
        pdf = PdfFileReader(pdf_ubicacion)
        doc = fitz.open(pdf_ubicacion)
    os.remove(str(lista_pdf_nuevo_unido_parte_1[0]))
    

#Insertar el logo y la firma en  los nuevos pdf
def insertar_logo_firma_Pdf():
    ruta_logo_ubicaion = ruta_logo[0]
    ruta_firma_ubicacion = ruta_firma[0]
    packet = io.BytesIO()
   
    can = canvas.Canvas(packet, pagesize=letter)
    can.drawImage(ruta_logo_ubicaion,500,760,70,70)
    can.drawImage(ruta_firma_ubicacion,10,23,145,65)
    can.save()

    packet.seek(0)
    new_pdf = PdfFileReader(packet)

    pdf_ubicacion = os.path.join(str(lista_pdf_nuevo_unido_parte_1[0]))
    pdf = PdfFileReader(pdf_ubicacion)
   
    for item in lista_pdf_individual:
        existing_pdf = PdfFileReader(item, "rb")
        pdfWriter = PdfFileWriter()
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        pdfWriter.addPage(page)
        outputStream = open(item, "wb")
        pdfWriter.write(outputStream)
        outputStream.close()

        pdf_ubicacion = os.path.join(item)
        pdf = PdfFileReader(pdf_ubicacion)
        doc = fitz.open(pdf_ubicacion)
    insertar_neto_pdf()


def devolverArchivos(folder):
	for archivo in os.listdir(folder):
		lista_pdf_plame.append(os.path.join(folder,archivo))
		if os.path.isdir(os.path.join(folder,archivo)):
			devolverArchivos(os.path.join(folder,archivo))   

#Unir  los pdf de la carpeta principal y generar uno solo     
def unir_archivos_antes(folder):
    fusionador = PdfFileMerger()
    tes1 = len(lista_pdf_plame)
    posision1 = tes1 -1
    lista_pdf_plame.pop(int(posision1))
    lista_pdf_nuevo_unido_parte_1.append(folder+"/Boletas_PLAME/BoletasUnidas.pdf")
    #print(lista_pdf_nuevo_unido_parte_1)
    for pdf in lista_pdf_plame:
        fusionador.append(open(pdf, 'rb'))
    with open(folder+"/Boletas_PLAME/BoletasUnidas.pdf", 'wb') as salida:
        fusionador.write(salida)


def separa_renombrar_pdf(folder):

    carpeta_destino = os.mkdir(os.path.join(folder,"Boletas") +'_PLAME') 
    carpeta = os.path.join(folder,"Boletas") +'_PLAME'
    devolverArchivos(folder)
    unir_archivos_antes(folder) 

    pdf_ubicacion = os.path.join(str(lista_pdf_nuevo_unido_parte_1[0]))
    pdf = PdfFileReader(pdf_ubicacion)
    doc = fitz.open(pdf_ubicacion) 

    for numero_paginas in range(pdf.numPages):
        pdfWriter = PdfFileWriter()
        pdfWriter.addPage(pdf.getPage(numero_paginas))
    
        #Buscamos el DNI en los pdf
        pagina_lectura_DNI_CI = doc.loadPage(numero_paginas)  
        pagina_lectura_DNI_CI_texto = pagina_lectura_DNI_CI.getText("text")  
        ubicacion_DNI_CI = pagina_lectura_DNI_CI_texto.find("Situación")
        Nombre_pdf_individual = pagina_lectura_DNI_CI_texto[ubicacion_DNI_CI+13:ubicacion_DNI_CI+22]

        #Buscamos el adelanto de sueldo y el neto a pagar en el pdf
        adelanto_sueldo_lectura = doc.loadPage(numero_paginas)  
        lectura_texto = adelanto_sueldo_lectura.getText("text")  
        ubicacion_adelanto_sueldo = lectura_texto.find("ADELANTO")
        ubicacion_neto_pagar = lectura_texto.find("Neto a Pagar")
        adelato_sueldo = lectura_texto[ubicacion_adelanto_sueldo+9:ubicacion_adelanto_sueldo+16] 
        neto = lectura_texto[ubicacion_neto_pagar+12:ubicacion_neto_pagar+20] 

        #Si encuentra valor nulo en el adelanto , lo reemplaza por 0
        try:
            neto_pagado.update({str(Nombre_pdf_individual.strip())+".pdf" : round((float(adelato_sueldo.strip()) +float(neto.strip())),2) } )    
        except:
            try:
                print(Nombre_pdf_individual)
                string_1 = adelato_sueldo
                string_2 = neto
                string_1 = ''.join( x for x in string_1 if x not in characters)
                string_2 = ''.join( x for x in string_2 if x not in characters)
                neto_pagado.update({str(Nombre_pdf_individual.strip())+".pdf" : round((float(string_1.strip()) +float(string_2.strip())),2) } ) 
            except:
                adelanto_vacio = "0"
                neto_vacio = neto 
                neto_vacio = ''.join( x for x in string_2 if x not in characters)
                neto_pagado.update({str(Nombre_pdf_individual.strip())+".pdf" : round((float(adelanto_vacio.strip()) +float(neto_vacio.strip())),2) } ) 

        with open (os.path.join( os.path.join(folder,"Boletas") +'_PLAME','{0}.pdf'.format(Nombre_pdf_individual.strip(),numero_paginas)), "wb") as f:   
            lista_pdf_individual.append((folder + '/Boletas_PLAME/'+'{0}.pdf').format(Nombre_pdf_individual.strip()))
            pdfWriter.write(f)
            f.close()

def generar_pdf():
    abrir_archivo()

def generar_pdf_final():
    folder = ruta_folder[0]
    separa_renombrar_pdf(folder)
    insertar_logo_firma_Pdf()
    armado_pdf_final()

def insertarlogo():
    conexion_bd()
    abrir_archivo_logo();

def insertarFirma():
    abrir_archivo_firma()    


lista_pdf_plame=[]
lista_pdf_individual=[]
lista_pdf_nuevo_unido_parte_1 = []
lista_de_documento_ordenado = []
ruta_logo = []
ruta_firma = []
neto_pagado = {}
lista_unidades = []
lista_unidades_vista = []
lista_unidades_codigo = []
Undiad_Key = {}
ruta_folder = []
ruta_pdf = []
characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

#Creamos la ventana
v0 = tk.Tk()
v0.geometry("260x250")
v0.title('APLICATIVO DE PDF V1')
#centrar 
windowWidth = v0.winfo_reqwidth()
windowHeight = v0.winfo_reqheight()
positionRight = int(v0.winfo_screenwidth()/2 - windowWidth/2)
positionDown = int(v0.winfo_screenheight()/2 - windowHeight/2)
v0.geometry("+{}+{}".format(positionRight, positionDown))
#creacion de los com
Label(v0,text="Seleccione el logo de la empresa").place(x=20,y=5)
Button(v0,text="Seleccionar Logo",command=insertarlogo).place(x=20,y=30)
Label(v0,text="Seleccione la Imagen de La Firma").place(x=20,y=65)
Button(v0,text="Seleccionar Firma Digital",command=insertarFirma).place(x=20,y=90)
Label(v0,text="Seleccione archivo PDF").place(x=20,y=120)
Button(v0,text="Abrir archivo",command=generar_pdf).place(x=20,y=145)
Label(v0,text="Codigo Unidad :").place(x=20,y=175)
txt_cod_unidad = tk.Entry(v0)
txt_cod_unidad.place(x=120,y=175)
Button(v0,text="Generar Boletas",command=generar_pdf_final).place(x=20,y=200)
v0.mainloop()

'''
python -m pip install --upgrade pip
pip install PyPDF2
pip install fitz
pip install reportlab
pip install PyMuPDF
pip install pyodbc
'''