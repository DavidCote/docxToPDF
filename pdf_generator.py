import os
from docxtpl import DocxTemplate
from docx2pdf import convert

# Ruta de la plantilla .docx
# La ruta de guardado por defecto es la carpeta 'Converted' en la carpeta raiz del programa

'''
 template_path = ruta de la plantilla
 Cambia la ruta de la plantilla .docx
'''


def def_template(template_path):
    global template
    if os.path.exists(template_path):
        gpath, extension = os.path.splitext(template_path)
        if extension == ".docx":
            template = DocxTemplate(template_path)
        else:
            raise TypeError
    else:
        raise FileNotFoundError


'''
 doc_name = ruta de guardado del archivo generado docx
 values = diccionario con los valores a reemplazar en la plantilla
 Crea un archivo .docx a partir de una plantilla con los valores  
'''


def gen_docx(doc_name, values):
    template.render(values)
    template.save(doc_name)

'''
 source_doc = ruta del archivo .docx
 convierte la plantilla docx a un archivo .pdf 
'''


def gen_pdf(source_doc):
    if os.path.exists(source_doc):
        convert(source_doc)
    else:
        raise FileNotFoundError
