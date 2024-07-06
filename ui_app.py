import os
import csv_reader
import pdf_generator
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


# Clase Application hereda de Frame, contruye la interfaz de usuario


class Application(tk.Frame):

    def __init__(self, window, menu_bar):
        super().__init__(window)
        self.default_route = os.path.join(os.getcwd(), "Converted")  # Ruta por defecto de guardado
        self.window = window
        self.menu_bar = menu_bar

        self.window.title('Generador')  # Titulo de la venta principal del programa
        self.window.iconbitmap('6E1.ico')  # Icono que aparece en todas las ventanas del programa
        self.window.resizable(False, False)  # Impide que la ventana del programa cambie de tamaño

        self.build_menu()
        self.build_widgets()

        # Los objetos creados dentro de la venta ocupan todo el espacio disponible

        self.pack(fill='both', expand=True)

    # Metodo que contruye la barra de menu

    def build_menu(self):
        # Menus
        self.menu_inicio = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_ayuda = tk.Menu(self.menu_bar, tearoff=0)

        # Opciones dentro de los menus

        self.menu_bar.add_cascade(label='Inicio', menu=self.menu_inicio)
        self.menu_bar.add_cascade(label="Ayuda", menu=self.menu_ayuda)

        self.menu_ayuda.add_cascade(label="Acerca de...")
        self.menu_ayuda.add_cascade(label="Licencia")

        # Funcionalidad de opciones

        self.menu_inicio.add_command(label='Generar todo', command=self.gen_all)
        self.menu_inicio.add_command(label="Generar plantillas", command=self.gen_doc)
        self.menu_inicio.add_separator()
        self.menu_inicio.add_command(label='Salir', command=self.window.destroy)

    # Metodo que construye todos los elementos de la interfaz (texto, entradas y botones)

    def build_widgets(self):
        # Etiquetas de cada campo

        self.label_csv = tk.Label(self, text='Archivo .csv')
        self.label_csv.grid(row=0, column=0, columnspan=2, sticky="w", padx=(20, 3), pady=(20, 0))

        self.label_plantilla = tk.Label(self, text='Plantilla .docx')
        self.label_plantilla.grid(row=2, column=0, columnspan=2, sticky="w", padx=(20, 3), pady=(20, 0))

        self.label_guardar = tk.Label(self, text='Carpeta de guardado')
        self.label_guardar.grid(row=4, column=0, columnspan=2, sticky="w", padx=(20, 3), pady=(20, 0))

        # Entradas de cada campo

        self.var_csv = tk.StringVar()
        self.entry_csv = tk.Entry(self, textvariable=self.var_csv)
        self.entry_csv.config()
        self.entry_csv.grid(row=1, column=0, columnspan=2, sticky='EW', padx=20, pady=(0, 0))

        self.var_plantilla = tk.StringVar()
        self.entry_plantilla = tk.Entry(self, textvariable=self.var_plantilla)
        # self.entry_plantilla.config(columnspan)
        self.entry_plantilla.grid(row=3, column=0, columnspan=2, sticky='EW', padx=20, pady=(0, 0))

        self.var_guardado = tk.StringVar()
        self.entry_guardado = tk.Entry(self, textvariable=self.var_guardado)
        # self.entry_guardado.config(columnspan)
        self.entry_guardado.grid(row=5, column=0, columnspan=2, sticky='EW', padx=20, pady=(0, 0))

        # Botones

        self.boton_examinar = tk.Button(self, text='Examinar...', bd=1, command=self.open_csv)
        self.boton_examinar.config(width=15)
        self.boton_examinar.grid(row=1, column=2, padx=(0, 20), pady=10)

        self.boton_examinar2 = tk.Button(self, text='Examinar...', bd=1, command=self.open_docx)
        self.boton_examinar2.config(width=15)
        self.boton_examinar2.grid(row=3, column=2, padx=(0, 20), pady=10)

        self.boton_cambiar = tk.Button(self, text='Cambiar', bd=1, command=self.ch_directory)
        self.boton_cambiar.config(width=15)
        self.boton_cambiar.grid(row=5, column=2, padx=(0, 20), pady=10)

        self.boton_generar = tk.Button(self, text='Generar Todo', bd=1, command=self.gen_all)
        self.boton_generar.config(width=20)
        self.boton_generar.grid(row=6, column=0, padx=20, pady=(20, 20))

        self.boton_convertir = tk.Button(self, text='Generar plantillas', bd=1, command=self.gen_doc)
        self.boton_convertir.config(width=20)
        self.boton_convertir.grid(row=6, column=1, padx=20, pady=(20, 20))

    # Metodo asociado a boton_examinar, obtiene la ruta del archivo .csv y la escribe en su entrada

    def open_csv(self):
        file_path = os.path.normpath(
            filedialog.askopenfilename(filetypes=(("csv files", "*.csv"), ("all files", "*.*"))))
        self.var_csv.set(file_path)

    # Metodo asociado a boton_examinar2, obtiene la ruta de la plantilla .docx y la escribe en su entrada

    def open_docx(self):
        file_path = os.path.normpath(
            filedialog.askopenfilename(filetypes=(("docx files", "*.docx"), ("all files", "*.*"))))
        self.var_plantilla.set(file_path)

    # Metodo asociado a boton_cambiar, obtiene la ruta de la carpeta de guardado y la escribe en su entrada

    def ch_directory(self):
        dir_path = os.path.normpath(filedialog.askdirectory())
        self.var_guardado.set(dir_path)

    '''
     Metodo asociado a boton_generar, genera los archivos .docx y los conviete a .pdf; 
     guarda en una lista los diccionarios obtenidos del archivo .csv, verifica
     si hay una ruta de guardado seleccionada, sino, establece una ruta predefinida.
     Verifica que la plantilla exista y sea del tipo requerido (.docx)
     Dentro del bucle se extraen los valores del primer campo del diccionario junto con su indice, 
     estos dos  valores se usan para crear el nombre del nuevo archivo .docx que tambien es usado para crear el .pdf
     finalmente una bandera indica si el proceso fue terminado sin ningun error
    '''

    def gen_all(self):
        flag = False
        ls = csv_reader.read_csv(self.var_csv.get())

        if self.var_guardado.get() == "" and (not os.path.exists(self.default_route)):
            os.mkdir(self.default_route)
            self.var_guardado.set(self.default_route)
        elif self.var_guardado.get() == "" and os.path.exists(self.default_route):
            self.var_guardado.set(self.default_route)

        try:
            pdf_generator.def_template(self.var_plantilla.get())
        except FileNotFoundError:
            messagebox.showerror("Archivo no encontrado", "La plantilla especificada no existe")
        except TypeError:
            messagebox.showerror("Tipo de archivo", "La plantilla debe ser un documento de Word (.docx)")
        else:
            flag = True

        for row_dicc in ls:
            temp = list(row_dicc.values())
            doc_name = f"{temp[0]}_{ls.index(row_dicc) + 1}.docx"
            save_path = os.path.join(self.var_guardado.get(), doc_name)
            try:
                pdf_generator.gen_docx(save_path, row_dicc)
            except NameError:
                print("Imposible generar nuevos docx")
                break
            else:
                flag = True
            try:
                pdf_generator.gen_pdf(save_path)
            except FileNotFoundError:
                print("no hay archivo docx y no se pudo generar pdf")
                break
            else:
                flag = True

        if flag:
            messagebox.showinfo("Proceso terminado con exito", f"Se generaron y convirtieron  {len(ls)} plantillas")

    '''
     Metodo asociado a boton_generar, genera unicamente los archivos .docx; 
     guarda en una lista los diccionarios obtenidos del archivo .csv, verifica
     si hay una ruta de guardado seleccionada, sino, establece una ruta predefinida.
     Verifica que la plantilla exista y sea del tipo requerido (.docx)
     Dentro del bucle se extraen los valores del primer campo del diccionario junto con su indice, 
     estos dos  valores se usan para crear el nombre del nuevo archivo .docx 
     finalmente una bandera indica si el proceso fue terminado sin ningun error
    '''

    def gen_doc(self):
        flag = False
        ls = csv_reader.read_csv(self.var_csv.get())

        if self.var_guardado.get() == "" and (not os.path.exists(self.default_route)):
            os.mkdir(self.default_route)
            self.var_guardado.set(self.default_route)
        elif self.var_guardado.get() == "" and os.path.exists(self.default_route):
            self.var_guardado.set(self.default_route)

        try:
            pdf_generator.def_template(self.var_plantilla.get())
        except FileNotFoundError:
            messagebox.showerror("Archivo no encontrado", "La plantilla especificada no existe")
        except TypeError:
            messagebox.showerror("Tipo de archivo", "La plantilla debe ser un documento de Word (.docx)")
        else:
            flag = True

        for row_dicc in ls:
            temp = list(row_dicc.values())
            doc_name = f"{temp[0]}_{ls.index(row_dicc) + 1}.docx"
            save_path = os.path.join(self.var_guardado.get(), doc_name)
            try:
                pdf_generator.gen_docx(save_path, row_dicc)
            except NameError:
                print("Imposible generar nuevos docx")
                break
            else:
                flag = True

        if flag:
            messagebox.showinfo("Proceso terminado con exito", f"Se generaron {len(ls)} plantillas")
