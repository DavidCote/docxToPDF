import csv
from  tkinter import  messagebox

'''
file_path = ruta del archivo .csv
El archivo se lee, se garda en forma de diccionario y se retorna una lista
'''


def read_csv(file_path):
    lista = list()
    try:
        with open(file_path, newline='') as file:
            reader = csv.DictReader(file)
            for r in reader:
                lista.append(r)
    except FileNotFoundError:
        messagebox.showerror("Archivo no encontrado", "El archivo .csv especificada no existe")
    except UnicodeDecodeError:
        messagebox.showerror("Error de lectura", "Es imposible leer el archivo")
    return lista
