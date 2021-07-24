import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import cv2
import imutils
from tkinter import *
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import ImageTk, Image

root = tk.Tk()
root.geometry("1280x736")
root.title("PSQ-APP")
root.resizable(width=False, height=False)

fondo = tk.PhotoImage(file="fondo.png")
fondo1 = tk.Label(root, image=fondo).place(x=0, y=0)







nombre1,apellido1,edad1,correo1,telefono1 = [],[],[],[],[]

def agregar_datos():
	global nombre1, apellido1, dni1, correo1, telefono1

	nombre1.append(ingresa_nombre.get())
	apellido1.append(ingresa_apellido.get())
	edad1.append(ingresa_edad.get())
	correo1.append(ingresa_correo.get())
	telefono1.append(ingresa_telefono.get())

	ingresa_nombre.delete(0,END)
	ingresa_apellido.delete(0,END)
	ingresa_edad.delete(0,END)
	ingresa_correo.delete(0,END)
	ingresa_telefono.delete(0,END)





	datos = {'DIA':nombre1,'T1':apellido1, 'T2':edad1, 'T3':correo1, 'T4':telefono1 }
	nom_excel  = str("datatime.xlsx")
	df = pd.DataFrame(datos,columns =  ['DIA', 'T1', 'T2', 'T3', 'T4'])
	df.to_excel(nom_excel)

def grafico():
    global exl1, df, valores, gp, canvas, df1, ax1, bar1

    exl1 = "datatime.xlsx"
    df = pd.read_excel(exl1)

    valores = df[["DIA","T4"]]

    df1 = valores.plot.bar(x="DIA", y="T4", rot = 10)
    plt.show()


    figure1 = plt.Figure(figsize=(4.8, 4.3), dpi=100)
    ax1 = figure1.add_subplot(111)
    bar1 = FigureCanvasTkAgg(figure1, root)
    bar1.get_tk_widget().place(x=780, y=300)
    df1 = valores[['DIA','T4']].groupby('DIA').sum()
    df1.plot(kind='bar', legend=True, ax=ax1)
    ax1.set_title('Ultimo tiempo por dia')



def iniciar():
    global cap
    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
    visualizar()


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/Desktop",
                                          title="Seleccionar DATATIME",
                                          filetype=(("xlsx files", "datatime.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == "datos.xlsx":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row)
    return None


def clear_data():
    tv1.delete(*tv1.get_children())
    return None


def clear_data():
    tv1.delete(*tv1.get_children())
    return None

def visualizar():
    global cap
    if cap is not None:
        ret, frame = cap.read()
        if ret == True:
            frame = imutils.resize(frame, width=580)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            im = Image.fromarray(frame)
            img = ImageTk.PhotoImage(image=im)
            lblVideo.configure(image=img)
            lblVideo.image = img
            lblVideo.after(10, visualizar)
        else:
            lblVideo.image = ""
            cap.release()


def finalizar():
    global cap
    cap.release()


cap = None
btnIniciar1 = Button(root, text="Ampliar", width=10, command=iniciar)
btnIniciar1.place(x=650, y=370)
btnIniciar = Button(root, text="Iniciar", width=10, command=iniciar)
btnIniciar.place(x=650, y=470)
btnFinalizar = Button(root, text="Finalizar", width=10, command=finalizar)
btnFinalizar.place(x=650, y=500)
lblVideo = Label(root)
lblVideo.place( x=30, y=95)

file_frame = tk.LabelFrame(root, text="Ingresar datos de tiempo", bg='MediumPurple4', fg='white')
file_frame.place(height=180, width=480, rely=0.13, relx=0.61)

frame1 = tk.LabelFrame(root, text="Excel Data", bg='MediumPurple4', fg='white')
frame1.place(height=170, width=730, rely=0.75, relx=0.02)


nombre = Label(root, text ='DIA',  bg='indigo', fg='white').place(rely=0.16, relx=0.65)
ingresa_nombre = Entry(root,  width=13, font = ('Arial',8))
ingresa_nombre.place(rely=0.16, relx=0.67)

apellido = Label(root, text ='T1', bg='indigo', fg='white').place(rely=0.20, relx=0.65)
ingresa_apellido = Entry(root, width=13, font = ('Arial',8))
ingresa_apellido.place(rely=0.20, relx=0.67)

edad = Label(root, text ='T2', bg='indigo', fg='white').place(rely=0.24, relx=0.65)
ingresa_edad = Entry(root,  width=13, font = ('Arial',8))
ingresa_edad.place(rely=0.24, relx=0.67)

correo = Label(root, text ='T3', bg='indigo', fg='white').place(rely=0.28, relx=0.65)
ingresa_correo = Entry(root,  width=13, font = ('Arial',8))
ingresa_correo.place(rely=0.28, relx=0.67)

telefono = Label(root, text ='T4', bg='indigo', fg='white').place(rely=0.32, relx=0.65)
ingresa_telefono = Entry(root, width=13, font = ('Arial',8))
ingresa_telefono.place(rely=0.32, relx=0.67)

agregar = Button(root, width=7, text='Guardar',fg= 'white', bg='indigo',bd=5, command =agregar_datos)
agregar.place(rely=0.16, relx=0.80)
"""
guardar = Button(root, width=7, text='Guardar',fg= 'white', bg='indigo',bd=5, command =guardar_datos)
guardar.place(rely=0.91, relx=0.80)
"""

# Buttons
button1 = tk.Button(file_frame, text="Buscar Archivo", bg='mediumpurple', fg='white', command=lambda: File_dialog())
button1.place(rely=0.50, relx=0.68)

button2 = tk.Button(file_frame, text="Cargar Datos" ,bg='mediumpurple', fg='white', command=lambda: Load_excel_data())
button2.place(rely=0.50, relx=0.48)

button3 = tk.Button(file_frame, text="Graficar" ,bg='mediumpurple', fg='white', command=lambda: grafico())
button3.place(rely=0.70, relx=0.48)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0.93, relx=0)


## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget


root.mainloop()
