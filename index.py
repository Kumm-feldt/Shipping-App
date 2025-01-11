from tkinter import *
from PIL import Image, ImageTk
from pathlib import Path
from docxtpl import DocxTemplate
from tkinter import messagebox as MessageBox
import shutil
import pyglet, os
from datetime import date
import asyncio


today = date.today()
c_date = today.strftime("%d/%m/%Y")

path = Path()


# DIRECTORIOS

directorio2 = "C:\\Users\\antho\\OneDrive\\Desktop\\"


doc = ["solicitud-zarpe.docx", "cartas-de-flete.docx", "solicitud-visita.docx", "solicitud-ingreso-personal.docx", "solicitud-almacenamiento-patio-bodega.docx", "declaracion-oficial-arribo-buque.docx"]


ruta_actual = os.path.abspath(os.getcwd())

f = [
    "\\carga-general.xlsx",
    "\\contenedores.xlsx",
    "\\contenedores-vacios.xlsx",
    "\\declaracion-oficial-arribo-buque.docx",
    "\\formato-precalculo.xlsx",
    "\\formato-reporte-aribo.xlsx",
    "\\libretin-descarga.xlsx",
    "\\port-log.xlsx",
    "\\registro-patio-bodega.xlsx",
    "\\solicitud-almacenamiento-patio-bodega.docx",
    "\\cartas-de-flete.docx",
    "\\solicitud-ingreso-personal.docx",
    "\\solicitud-visita.docx",
    "\\solicitud-zarpe.docx"
]

#-------------------------------------
#Fonts
pyglet.font.add_file("Black.ttf")
pyglet.font.add_file("Medium.ttf")
pyglet.font.add_file("Regular.ttf")
pyglet.font.add_file("Light.ttf")
pyglet.font.add_file("Thin.ttf")

pyglet.font.load(name='Arial', size=12)

#pyglet.font.load("Black.ttf")
pyglet.font.add_file("Medium.ttf")
pyglet.font.add_file("Regular.ttf")
pyglet.font.add_file("Light.ttf")
pyglet.font.add_file("Thin.ttf")

#------------ root -------------------
root = Tk()
root.geometry("1350x800")
root.configure(bg='white')

root.title("GUATEMALA MARITIMA, S.A.")

#---------------- Images---------------
my_img = Image.open("gt.png")
my_img = my_img.resize((824,800), Image.Resampling.LANCZOS)
img = ImageTk.PhotoImage(my_img)
label_img = Label(root, image = img)
label_img.place(x=11, y=0)
label_img.configure(bg='white')


title = Image.open("title.png")
title = title.resize((297,145), Image.Resampling.LANCZOS)
img1 = ImageTk.PhotoImage(title)
label_img1 = Label(root, image = img1)
label_img1.place(x=915, y=139)
label_img1.configure(bg='white')

#--Flecha
arrow = Image.open("flecha.png")
arrow = arrow.resize((50,49), Image.Resampling.LANCZOS)
img2 = ImageTk.PhotoImage(arrow)
label_img2 = Label(root, image = img2)
label_img2.place(x=515, y=696)
label_img2.configure(bg='white', fg="white")

arrow2 = Image.open("flecha.png")
arrow2 = arrow2.resize((25,24), Image.Resampling.LANCZOS)
img_arrow2 = ImageTk.PhotoImage(arrow2)
label_img_a = Label(root, image = img_arrow2)
label_img_a.configure(bg='white', fg="white")
#--------------------------------------------

rounded = Image.open("rounded_button.png")
rounded = rounded.resize((181,48), Image.Resampling.LANCZOS)
img3 = ImageTk.PhotoImage(rounded)

#--------------- Text -------------------

psj=Label(root, text="P u e r t o  S a n  J o s e,  E s c u i n t l a", font=("Heebo Light", 15), fg="#000000",bg="white")
psj.place(x=497, y=10)

#FRAMES
btm_frame = Frame(root, bg='#a6a2a2', width=1000, height=250, pady=3)
#----
port1_frame = Frame(btm_frame, bg='#ffffff', width=165, height=100, padx=3, pady=3)
port2_frame = Frame(btm_frame, bg='#ffffff', width=165, height=100, padx=3, pady=3)
port3_frame = Frame(btm_frame, bg='#ffffff', width=165, height=100, padx=3, pady=3)
port4_frame = Frame(btm_frame, bg='#ffffff', width=165, height=100, padx=3, pady=3)
port5_frame = Frame(btm_frame, bg='#ffffff', width=165, height=100, padx=3, pady=3)

#Variables Globales de Funciones
tituloNombre = Label(root, text="Nombre del Barco", font=("Heebo Medium", 18), fg="#000000",
                bg="white")



# ---------------------- TITULOS FORMULARIO ------------------
nombreBarco_t = Label(root, text="", font=("Heebo Medium", 12), bg="white", fg="#5E5E5E")
#----------
agent_t = Label(root, text="Nombre del Agente (agent)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )

steevedoring_t = Label(root, text="Estivadora (steevedoring)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )

vessel_name_t = Label(root, text="Nombre del Barco", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
trip_t = Label(root, text="Viaje (trip)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E", anchor="e")
master_t = Label(root, text="Nombre del Capitan (master)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
flag_t = Label(root, text="Bandera (flag)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
imo_t = Label(root, text="I.M.O", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
construction_country_t = Label(root, text="Lugar de construccion", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
type_vessel_t = Label(root, text="Tipo de Embarcacion", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
call_sign_t = Label(root, text="Indicativo de Llamada (call sign)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
t_r_b_t = Label(root, text="T.R.B | G.R.T", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
t_r_n_t = Label(root, text="T.R.N | N.R.T", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
l_o_a_t = Label(root, text="Eslora (loa)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
beam_t = Label(root, text="Manga (beam)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
draft_t = Label(root, text="Calado (draft)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
owner_t = Label(root, text="Armadores u Operadores (owners)", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
port_register_t = Label(root, text="Puerto de Registro", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
date_register_t = Label(root, text="Fecha de Registro", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
port_t = Label(root, text="Puerto de Atraque", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
arrival_date_t = Label(root, text="Fecha de Arribo", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
last_port_t = Label(root, text="Ultimo Puerto", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
next_port_t = Label(root, text="Siguiente Puerto", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
crew_t = Label(root, text="Numero de Tripulantes", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
passengers_t = Label(root, text="Numero de Pasajeros", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
berthing_date_t = Label(root, text="Hora de Atraque", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
anchored_time_t = Label(root, text="Hora de Fondeo", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )

text= Label(root, text="INSERTA ULTIMOS PUERTOS:" )
port_5_t = Label(port5_frame, text="5. Puerto", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
c_5_t = Label(port5_frame, text="5. Pais", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
d_5_t = Label(port5_frame, text="5. Fecha", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )

port_4_t = Label(port4_frame, text="4. Puerto", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
c_4_t = Label(port4_frame, text="4. Pais", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
d_4_t = Label(port4_frame, text="4. Fecha", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )

port_3_t = Label(port3_frame, text="3. Puerto", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
c_3_t = Label(port3_frame, text="3. Pais", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
d_3_t = Label(port3_frame, text="3. Fecha", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )

port_2_t = Label(port2_frame, text="2. Puerto", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
c_2_t = Label(port2_frame, text="2. Pais", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
d_2_t = Label(port2_frame, text="2. Fecha", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )

port_1_t = Label(port1_frame, text="1. Puerto", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
c_1_t = Label(port1_frame, text="1. Pais", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )
d_1_t = Label(port1_frame, text="1. Fecha", font=("Heebo Medium", 11), bg="white", fg="#5E5E5E" )


#---------------------- VARIABLES FORMULARIO-------------------

nombreBarco = Entry(root, width=22, borderwidth=2, font=('Heebo Medium', 13))
vessel_name_e = nombreBarco.get()
#--------------
agent_e = Entry(root, width=21, font=("Heebo Medium", 12))
steevedoring_e = Entry(root, width=21, font=("Heebo Medium", 12))
trip_e = Entry(root, width=21, font=("Heebo Medium", 12))
master_e = Entry(root, width=21, font=("Heebo Medium", 12))
flag_e = Entry(root, width=21, font=("Heebo Medium", 12))
imo_e = Entry(root, width=21, font=("Heebo Medium", 12))
construction_country_e = Entry(root, width=21, font=("Heebo Medium", 12))
type_vessel_e = Entry(root,width=21, font=("Heebo Medium", 12))
call_sign_e = Entry(root, width=21, font=("Heebo Medium", 12))
t_r_b_e = Entry(root, width=21, font=("Heebo Medium", 12))
t_r_n_e = Entry(root, width=21, font=("Heebo Medium", 12))
l_o_a_e = Entry(root, width=21, font=("Heebo Medium", 12))
beam_e = Entry(root, width=21, font=("Heebo Medium", 12))
draft_e = Entry(root, width=21, font=("Heebo Medium", 12))
owner_e = Entry(root, width=21, font=("Heebo Medium", 12))
port_register_e = Entry(root, width=21, font=("Heebo Medium", 12))
date_register_e = Entry(root, width=21, font=("Heebo Medium", 12))
port_e = Entry(root, width=21, font=("Heebo Medium", 12))
arrival_date_e = Entry(root, width=21, font=("Heebo Medium", 12))
last_port_e = Entry(root,width=21, font=("Heebo Medium", 12))
next_port_e = Entry(root,width=21, font=("Heebo Medium", 12))
crew_e = Entry(root,width=21, font=("Heebo Medium", 12))
passengers_e = Entry(root, width=21, font=("Heebo Medium", 12))
berthing_date_e = Entry(root, width=21, font=("Heebo Medium", 12))
anchored_time_e = Entry(root, width=21, font=("Heebo Medium", 12))

port_5_e = Entry(port5_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
c_5_e = Entry(port5_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
d_5_e = Entry(port5_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)

port_4_e = Entry(port4_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
c_4_e = Entry(port4_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
d_4_e = Entry(port4_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)

port_3_e = Entry(port3_frame,width=18, font=("Heebo Medium", 11), borderwidth=3)
c_3_e = Entry(port3_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
d_3_e = Entry(port3_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)

port_2_e = Entry(port2_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
c_2_e = Entry(port2_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
d_2_e = Entry(port2_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)

port_1_e = Entry(port1_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
c_1_e = Entry(port1_frame, width=18, font=("Heebo Medium", 11), borderwidth=3)
d_1_e = Entry(port1_frame,width=18, font=("Heebo Medium", 11), borderwidth=3)




# Crear Carpeta y Excel

def next_section(): #siguiente >>>>
    button_new.destroy()
    button_ask.destroy()
    label_img1.place_forget()
    label_img2.destroy()
    tituloNombre.place(x=976, y=300)
    global nombreBarco
    nombreBarco.place(x=980, y=350)
    button_next.place(x=1020, y=400)
    label_img_a.place(x=1142, y=407)

def formSection(): #creacion de carpetas
    global name
    name=nombreBarco.get()

    # Crea Carpeta Principal
    fileName = f'Barco MV {name}'
    directorio = f"C:\\Users\\antho\\OneDrive\\Desktop\\{fileName}"
    os.mkdir(directorio)

    # Listado de Subcarpetas a Crear
    directory_list = ["BLs", "BOL. PAGO PRECALCULO", "CARTAS DE CORRECCION", "CARTAS DE DESPACHO", " FACTURAS EPQ", "SCANS DE BLs",
                      "SOL. ACTIVIDAD PERMITIDA", "JUST. SOBRANTES & FALTANTES", "DOCS. CAPITAN", "REC. REINTEGRO"]
    global paths
    for directories in directory_list:

        # Crea las SubCarpetas
        paths= os.path.join(directorio + "\\", directories)
        os.mkdir(paths)

        # Agrega ciertos documentos a ciertas SubCarpetas
        if directories == "SOL. ACTIVIDAD PERMITIDA":
            file = "\\solicitud-de-operacion-actividad-en-aduanas.xlsx"
            destino_final = f"{paths}{file}"
            insiderute = f"{ruta_actual}{file}"
            shutil.copy(insiderute, destino_final)

        elif directories == "JUST. SOBRANTES & FALTANTES":
            file = "\\SAT-justificaciones.docx"
            destino_final = f"{paths}{file}"
            insiderute = f"{ruta_actual}{file}"
            shutil.copy(insiderute, destino_final)

        elif directories == "REC. REINTEGRO":
            file = "\\recibo-caja-reintegro.xlsx"
            destino_final = f"{paths}{file}"
            insiderute = f"{ruta_actual}{file}"
            shutil.copy(insiderute, destino_final)

    # Copia Archivos Generales -> Unicamente .xlsx
    for file in f:
        destino_final = f'{directorio}{file}'
        caracter = destino_final[-5:]
        insiderute = f"{ruta_actual}{file}"
        if caracter == '.xlsx':
            shutil.copy(insiderute, destino_final)
            oldName = destino_final
            final_str = destino_final[:-5]
            newName= f'{final_str}-{name}{caracter}'
            os.rename(oldName, newName)

    root.geometry("1085x845")
    button_next.destroy()
    tituloNombre.destroy()
    nombreBarco.place(x=-1, y=-1)
    label_img.destroy()
    label_img_a.destroy()
    psj.destroy()
    button_exit.destroy()
    button_save.place(x=435, y=788)
    button_exit_2.place(x=1000, y=795)
    #global


    #Posicionando las entradas y titulos
    form_title=Label(root, text="Formulario de Buques Nuevos", font=("Heebo Medium", 14), bg="white", fg="black")
    form_title.grid(padx=10, pady=10,columnspan=10, column=0, row=0)
    agent_t.grid(column=0, row=1, padx=20,pady=5, sticky="w")
    agent_e.grid( column=0, row=2, padx=20, sticky= "n")
    steevedoring_t.grid( column=0, row=3, padx=20,pady=5, sticky="w")
    steevedoring_e.grid( column=0, row=4, padx=20, sticky= "n")
    trip_t.grid( column=0, row=5, padx=20,pady=5, sticky="w")
    trip_e.grid(column=0, row=6, padx=20, sticky= "n")
    master_t.grid(column=0, row=7, padx=20,pady=5, sticky="w")
    master_e.grid(column=0, row=8, padx=20, sticky= "n")
    flag_t.grid(column=0, row=9, padx=20,pady=5, sticky="w")
    flag_e.grid(column=0, row=10, padx=20, sticky= "n")
    imo_t.grid(column=0, row=11, padx=20,pady=5, sticky="w")
    imo_e.grid(column=0, row=12, padx=20, sticky= "n")
    construction_country_t.grid(column=0, row=13, padx=20,pady=5, sticky="w")
    construction_country_e.grid(column=0, row=14, padx=20, sticky= "n")


    # segunda columna
    call_sign_t.grid(pady=5, sticky="w", column=1, row=1, padx=30)
    call_sign_e.grid(column=1,sticky="w", row=2, padx=30)
    t_r_b_t.grid(pady=5, sticky="w", column=1, row=3, padx=30)
    t_r_b_e.grid(column=1, row=4,sticky="w", padx=30)
    t_r_n_t.grid(pady=5, sticky="w", column=1, row=5, padx=30)
    t_r_n_e.grid(column=1, row=6,sticky="w", padx=30)
    l_o_a_t.grid(pady=5, sticky="w", column=1, row=7, padx=30)
    l_o_a_e.grid(sticky="w", column=1, row=8, padx=30)
    beam_t.grid(pady=5, sticky="w", column=1, row=9, padx=30)
    beam_e.grid(sticky="w", column=1, row=10, padx=30)
    draft_t.grid(pady=5, sticky="w", column=1, row=11, padx=30)
    draft_e.grid(sticky="w", column=1, row=12, padx=30)
    type_vessel_t.grid(pady=5, sticky="w", column=1, row=13, padx=30)
    type_vessel_e.grid(column=1, row=14, padx=30, sticky="w")



    #tercera columna
    owner_t.grid(padx=30,pady=5, sticky="w", column=2, row=1)
    owner_e.grid(padx=30, sticky="w", column=2, row=2)
    port_register_t.grid(padx=30,pady=5, sticky="w", column=2, row=3)
    port_register_e.grid(padx=30, sticky="w", column=2, row=4)
    date_register_t.grid(padx=30,pady=5, sticky="w", column=2, row=5)
    date_register_e.grid(padx=30, sticky="w", column=2, row=6)
    port_t.grid(padx=30,pady=5, sticky="w", column=2, row=7)
    port_e.grid(padx=30, sticky="w", column=2, row=8)
    arrival_date_t.grid(padx=30,pady=5, sticky="w", column=2, row=9)
    arrival_date_e.grid(padx=30, sticky="w", column=2, row=10)
    last_port_t.grid(padx=30,pady=5, sticky="w", column=2, row=11)
    last_port_e.grid(padx=30, sticky="w", column=2, row=12)
    next_port_t.grid(padx=30,pady=5, sticky="w", column=2, row=13)
    next_port_e.grid(padx=30, sticky="w", column=2, row=14)


    #cuarta columna
    crew_t.grid(padx=30,pady=5, sticky="w", column=3, row=1)
    crew_e.grid(padx=30, sticky= "w", column=3, row=2)
    passengers_t.grid(padx=30,pady=5, sticky="w", column=3, row=3)
    passengers_e.grid(padx=30, sticky= "w", column=3, row=4)
    berthing_date_t.grid(padx=30,pady=5, sticky="w", column=3, row=5)
    berthing_date_e.grid(padx=30, sticky= "w", column=3, row=6)
    anchored_time_t.grid(padx=30,pady=5, sticky="w", column=3, row=7)
    anchored_time_e.grid(padx=30, sticky= "w", column=3, row=8)
    #FRAMES
    btm_frame.grid(padx=10, pady=10,columnspan=10, column=0, row=17)
        #place(x=35, y=535)
    port_title = Label(root, text="--------------------- INGRESA ULTIMOS PUERTOS VISITADOS ---------------------", font=("Heebo Medium", 12), bg="white", fg="black")
    port_title.grid(padx=10, pady=20,columnspan=10, column=0, row=15)

    #layout

    port5_frame.grid(row=0, column=0, padx=5)
    port4_frame.grid(row=0, column=1, padx=5)
    port3_frame.grid(row=0, column=2, padx=5)
    port2_frame.grid(row=0, column=3, padx=5)
    port1_frame.grid(row=0, column=4, padx=5)

    port_5_t.grid(row=0, column=0, padx=5, sticky="w")
    port_5_e.grid(row=1, column=0, padx=5, sticky="w")
    c_5_t.grid(row=2, column=0, padx=5, sticky="w")
    c_5_e.grid(row=3, column=0, padx=5, sticky="w")
    d_5_t.grid(row=4, column=0, padx=5, sticky="w")
    d_5_e.grid(row=5, column=0, padx=5, sticky="w")

    port_4_t.grid(row=0, column=0, padx=5, sticky="w")
    port_4_e.grid(row=1, column=0, padx=5, sticky="w")
    c_4_t.grid(row=2, column=0, padx=5, sticky="w")
    c_4_e.grid(row=3, column=0, padx=5, sticky="w")
    d_4_t.grid(row=4, column=0, padx=5, sticky="w")
    d_4_e.grid(row=5, column=0, padx=5, sticky="w")

    port_3_t.grid(row=0, column=0, padx=5, sticky="w")
    port_3_e.grid(row=1, column=0, padx=5, sticky="w")
    c_3_t.grid(row=2, column=0, padx=5, sticky="w")
    c_3_e.grid(row=3, column=0, padx=5, sticky="w")
    d_3_t.grid(row=4, column=0, padx=5, sticky="w")
    d_3_e.grid(row=5, column=0, padx=5, sticky="w")

    port_2_t.grid(row=0, column=0, padx=5, sticky="w")
    port_2_e.grid(row=1, column=0, padx=5, sticky="w")
    c_2_t.grid(row=2, column=0, padx=5, sticky="w")
    c_2_e.grid(row=3, column=0, padx=5, sticky="w")
    d_2_t.grid(row=4, column=0, padx=5, sticky="w")
    d_2_e.grid(row=5, column=0, padx=5, sticky="w")

    port_1_t.grid(row=0, column=0, padx=5, sticky="w")
    port_1_e.grid(row=1, column=0, padx=5, sticky="w")
    c_1_t.grid(row=2, column=0, padx=5, sticky="w")
    c_1_e.grid(row=3, column=0, padx=5, sticky="w")
    d_1_t.grid(row=4, column=0, padx=5, sticky="w")
    d_1_e.grid(row=5, column=0, padx=5, sticky="w")


#------------------------ gets--------------
agent_get = agent_e.get()
steevedoring_get = steevedoring_e.get()
trip_get = trip_e.get()
master_get = master_e.get()
flag_get = flag_e.get()
imo_get = imo_e.get()
construction_country_get = construction_country_e.get()
type_vessel_get = type_vessel_e.get()
call_sign_get = call_sign_e.get()
t_r_b_get = t_r_b_e.get()
t_r_n_get = t_r_n_e.get()
l_o_a_get = l_o_a_e.get()
beam_get = beam_e.get()
draft_get = draft_e.get()
owner_get = owner_e.get()
port_register_get = port_register_e.get()
date_register_get = date_register_e.get()
port_get = port_e.get()
arrival_date_get = arrival_date_e.get()
last_port_get = last_port_e.get()
next_port_get = next_port_e.get()
crew_get = crew_e.get()
passengers_get = passengers_e.get()
berthing_date_get = berthing_date_e.get()
anchored_time_get = anchored_time_e.get()
port_5_get = port_5_e.get()
c_5_get = c_5_e.get()
d_5_get = d_5_e.get()

port_4_get = port_4_e.get()
c_4_get = c_4_e.get()
d_4_get = d_4_e.get()

port_3_get = port_3_e.get()
c_3_get = c_3_e.get()
d_3_get = d_3_e.get()

port_2_get = port_2_e.get()
c_2_get = c_2_e.get()
d_2_get = d_2_e.get()

port_1_get = port_1_e.get()
c_1_get = c_1_e.get()
d_1_get = d_1_e.get()


#-------------------------------------------------------------------

def test():
    global ruta_actual
    context = {
        "c_date": c_date,
        'agent': agent_e.get(),
        'steevedoring': steevedoring_e.get(),
        'vessel_name': nombreBarco.get(),
        'trip': trip_e.get(),
        'master': master_e.get(),
        'flag': flag_e.get(),
        'imo': imo_e.get(),
        'construction_country': construction_country_e.get(),
        'type_vessel': type_vessel_e.get(),
        'call_sign': call_sign_e.get(),
        't_r_b': t_r_b_e.get(),
        't_r_n': t_r_n_e.get(),
        'l_o_a': l_o_a_e.get(),
        'beam': beam_e.get(),
        'draft': draft_e.get(),
        'owner': owner_e.get(),
        'port': port_e.get(),
        'arrival_date': arrival_date_e.get(),
        'last_port': last_port_e.get(),
        'next_port': next_port_e.get(),
        'crew': crew_e.get(),
        'port_register': port_register_e.get(),
        'date_register': date_register_e.get(),
        'passengers': passengers_e.get(),
        'berthing_date': berthing_date_e.get(),
        'anchored_time': anchored_time_e.get(),
        'port_5': port_5_e.get(),
        'c_5': c_5_e.get(),
        'd_5': d_5_e.get(),
        'port_4': port_4_e.get(),
        'c_4': c_4_e.get(),
        'd_4': d_4_e.get(),
        'port_3': port_3_e.get(),
        'c_3': c_3_e.get(),
        'd_3': d_3_e.get(),
        'port_2': port_2_e.get(),
        'c_2': c_2_e.get(),
        'd_2': d_2_e.get(),
        'port_1': port_1_e.get(),
        'c_1': c_1_e.get(),
        'd_1': d_1_e.get()
    }

    fileName = f'Barco MV {name}'
    directorio = f"C:\\Users\\antho\\Desktop\\{fileName}\\"
    caracter = '.docx'
    for documents in doc:
        # ../Desktop/Barco MV {name}
        destino_final = f'{directorio}{documents}'

        # Nombre del documento
        final_str2 = documents[:-5]

        # Nuevo Nombre del Documento
        newName_doc = f'{directorio}{final_str2}-{name}{caracter}' # .../Desktop/Barco Mv {ship}/document-{ship}.docx

        name_save = f'{final_str2}-{name}{caracter}'  # documento-{ship}.docx


        # Render document and save it as #documento-{ship}.docx
        document = DocxTemplate(documents)
        document.render(context)
        document.save(name_save)

        print(name_save)
        print(destino_final)
        # Mover documento de folder original hacia el destino final
        if(shutil.move(name_save, destino_final)):
            # Renombrar archivo a archivo + nombre de barco
            os.rename(f'{directorio}{documents}', newName_doc)
            print("renamed")


#FUNCIONES-------------------------->>>>>>>>>>>>>>>>>>>>>>>


def close():
    MessageBox.showinfo("GUARDADO", "Documentos Creados Correctamente")
    root.destroy()




#Botones
button_next = Button(root, text="SIGUIENTE", padx=18, pady=10, fg="#000000",bg="white",  font=("Heebo Medium", 10), borderwidth=0.1,command=lambda: formSection())
button_save = Button(root, image=img3 ,fg="#000000", bg="white", borderwidth=0, command= lambda:[test(),close()])
button_new = Button(root, text="Nuevo Barco", padx=50, pady=15, fg="#000000",bg="white",  font=("Heebo Medium", 12), borderwidth=0.5, command=lambda:next_section())
button_ask = Button(root, text="Consultar Barco", padx=39, pady=15, fg="#000000", bg="white",font=("Heebo Medium", 12),borderwidth=0.5)
button_exit = Button(root, text="CERRAR", padx=10, pady=5, command=quit, fg="#000000", bg="white", font=("Heebo Medium", 8), borderwidth=0)
button_exit_2 = Button(root, text="CERRAR", padx=10, pady=5, command=quit, fg="#000000", bg="white", font=("Heebo Medium", 8), borderwidth=0)

button_new.place(x=976, y=425)
button_ask.place(x=976, y=525)
button_exit.place(x=1042, y=745)


root.mainloop()