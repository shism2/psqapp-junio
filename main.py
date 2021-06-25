from tkinter import *
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from PIL import ImageTk, Image


root = Tk()
width= root.winfo_screenwidth()
height= root.winfo_screenheight()
#setting tkinter window size
root.geometry("%dx%d" % (width, height))
root.title('Guardar datos en Excel')


label = Label(root)
label.place(height=100, width=2000, rely=0, relx=0)
label.config(fg="snow",    # Foreground
             bg="whitesmoke",   # Background
             font=("Arial",24))

img = ImageTk.PhotoImage(Image.open("imagen.jpeg"))
panel = Label(root, image = img)
panel.place(height=100, rely=0, relx=0)


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


def guardar_datos():
	global nombre1,apellido1,edad1,correo1,telefono1

	datos = {'Nombre y Apellido':nombre1,'Nº Matricula':apellido1, 'Año Residencia':edad1, 'Institucion':correo1, 'Pais y Provincia':telefono1 }
	nom_excel  = str("datos.xlsx")
	df = pd.DataFrame(datos,columns =  ['Nombre y Apellido', 'Nº Matricula', 'Año Residencia', 'Institucion', 'Pais y Provincia'])
	df.to_excel(nom_excel)


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == "datos.csv":
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

file_frame = tk.LabelFrame(root, text="DATOS DEL ESTUDIANTE", bg='indigo', fg='white')
file_frame.place(height=200, width=2000, rely=0.1, relx=0)

nombre = Label(root, text ='Nombre y Apellido', bg='indigo', fg='white').place(rely=0.12, relx=0.02)
ingresa_nombre = Entry(root,  width=20, font = ('Arial',15))
ingresa_nombre.place(rely=0.12, relx=0.08)

apellido = Label(root, text ='Nº Matricula', width=10, bg='indigo', fg='white').place(rely=0.15, relx=0.02)
ingresa_apellido = Entry(root, width=20, font = ('Arial',15))
ingresa_apellido.place(rely=0.15, relx=0.08)

edad = Label(root, text ='Año Residencia', bg='indigo', fg='white').place(rely=0.13, relx=0.25)
ingresa_edad = Entry(root,  width=20, font = ('Arial',15))
ingresa_edad.place(rely=0.15, relx=0.25)

correo = Label(root, text ='Institucion', width=10, bg='indigo', fg='white').place(rely=0.19, relx=0.02)
ingresa_correo = Entry(root,  width=20, font = ('Arial',15))
ingresa_correo.place(rely=0.19, relx=0.08)

telefono = Label(root, text ='Pais y Provincia', bg='indigo', fg='white').place(rely=0.22, relx=0.02)
ingresa_telefono = Entry(root, width=20, font = ('Arial',15))
ingresa_telefono.place(rely=0.22, relx=0.08)


agregar = Button(root, width=20, text='Agregar',fg= 'white', bg='indigo',bd=5, command =agregar_datos)
agregar.place(rely=0.25, relx=0.10)
"""
archivo = Label(root, text ='Ingrese Nombre del archivo', width=25, bg='gray16',font = ('Arial',12, 'bold'), fg='white')
archivo.place(rely=0.12, relx=0.80)
"""
nombre_archivo = Entry(root, width=27, font = ('Arial',12),fg= 'white',highlightbackground = "indigo", highlightthickness=4)
nombre_archivo.place(rely=0.15, relx=0.80)

guardar = Button(root, width=20, text='Guardar',fg= 'white', bg='indigo',bd=5, command =guardar_datos)
guardar.place(rely=0.25, relx=0.26)


# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=200, width=750, rely=0.10, relx=0.37)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="VISUALIZAR DATOS", bg='mediumpurple', fg='white')
file_frame.place(height=100, width=400, rely=0.13, relx=0.78)

# Buttons
button1 = tk.Button(file_frame, text="Buscar Archivo", bg='mediumpurple', fg='white', command=lambda: File_dialog())
button1.place(rely=0.65, relx=0.50)

button2 = tk.Button(file_frame, text="Cargar Datos" ,bg='mediumpurple', fg='white', command=lambda: Load_excel_data())
button2.place(rely=0.65, relx=0.30)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)


## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget


root.mainloop()
