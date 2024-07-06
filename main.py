from tkinter import *
from ui_app import Application

'''
 Crea la venta principal del programa, crea la barra de menu y la asocia a la ventana principal
 finalmente pasa estos dos objetos al constructor de la Aplicacion y mantiene la ventana funcionando
'''


def main():
    window = Tk()
    menu_bar = Menu(window)
    window.config(menu=menu_bar)
    Application(window, menu_bar)
    window.mainloop()


if __name__ == '__main__':
    main()