import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
import re


#crear el libro de Excel

wb = Workbook()
ws = wb.active #para trabajar denreo de la hoja
ws.append(["Nombre", "Edad", "email", "Teléfono", "Dirección" ])
#wb.save('datos.xlsx') Crea el archivo datos con los campos nombre, edad,teléfono,dirección

#creo la función para guardar los datos

def guardar_datos():
    nombre = entry_nombre.get()
    edad = entry_edad.get()
    email = entry_email.get()
    telefono = entry_telefono.get()
    direccion = entry_direccion.get()

    if not nombre or not edad or not email or not telefono or not direccion:
        messagebox.showwarning(title="Advertencia", message="todos los campos son obligatorios")
        return
    try:
        edad = int(edad)
        telefono = int(telefono)
    except ValueError:
        messagebox.showwarning(title="Advertencia", message="Edad y teléfono deben ser números")
        return
    
    #validar formato email
    
    pattern = r"[^@]+@[^@]+\.[^@]+"

    if not re.match(pattern, email):
         messagebox.showwarning(title="Advertencia", message="El correo electrónico no es válido")

    
    ws.append = ([nombre,edad,email,telefono,direccion]) #tomo el valor de las variables
    wb.save('datos.xlsx') #Crea el archivo datos con los campos nombre, edad,teléfono,dirección
    messagebox.showinfo(title="información", message="Datos guardados con éxito")

    #borrando los campos

    entry_nombre.delete(0,tk.END)
    entry_edad.delete(0,tk.END)
    entry_telefono.delete(0 ,tk.END)
    entry_email.delete(0,tk.END)
    entry_direccion.delete(0,tk.END)


# crea la ventana de tkinter

root = tk.Tk()

#Título

root.title("Formulario de entrada de datos")

# dando color bg para background
root.configure(bg='#4B6587')

#creando estilos

label_style = {"bg": '#4B6587', "fg": "white"} # colro para etiquetas
entry_style = {"bg": '#D3D3D3', "fg": "black"} # color para entradas

               
# creando la primer etiqueta

label_nombre = tk.Label(root, text = "Nombre: ", **label_style)
label_nombre.grid(row=0, column=0, padx=10, pady=5)
entry_nombre = tk.Entry(root, **entry_style)
entry_nombre.grid(row=0, column=1, padx=10, pady=5)

#lo mismo para la edad

label_edad = tk.Label(root, text = "Edad: ", **label_style)
label_edad.grid(row=1, column=0, padx=10, pady=5)
entry_edad = tk.Entry(root, **entry_style)
entry_edad.grid(row=1, column=1, padx=10, pady=5)

# lo mismo para el email
label_email = tk.Label(root, text = "Email: ", **label_style)
label_email.grid(row=2, column=0, padx=10, pady=5)
entry_email = tk.Entry(root, **entry_style)
entry_email.grid(row=2, column=1, padx=10, pady=5)

# lo mismo para el teléfono


label_telefono = tk.Label(root, text = "Telefono: ", **label_style)
label_telefono.grid(row=3, column=0, padx=10, pady=5)
entry_telefono = tk.Entry(root, **entry_style)
entry_telefono.grid(row=3, column=1, padx=10, pady=5)

#lo mismo para la dirección


label_direccion = tk.Label(root, text = "Dirección: ", **label_style)
label_direccion.grid(row=4, column=0, padx=10, pady=5)
entry_direccion = tk.Entry(root, **entry_style) #root es la ventana
entry_direccion.grid(row=4, column=1, padx=10, pady=5)


#creo el botón para guardar

boton_guardar = tk.Button(root, text = "Guardar", command = guardar_datos, bg = '#6D8299', fg = 'white' , width=10) #root  para que este el boton en la ventana, el nombre de la funcion va sin paréntesis
boton_guardar.grid(row=5, column=0, columnspan= 2, padx=10, pady=5) #el columnspan es para 'centrar' el botón de guardar
                   




root.mainloop() #para que la ventana no se cierre al ejecutar





