import os
import csv_reader

path = r"c:\users\Coctel117\pycharmprojects\doc2pdf"
file = r"Template.docx"
ndir = "prueba"
empcsv = r"C:\Users\coctel117\PycharmProjects\doc2pdf\csvem.csv"
emwod = r"C:\Users\coctel117\PycharmProjects\doc2pdf\wem.docx"

ls = csv_reader.read_csv(r"C:\Users\coctel117\PycharmProjects\doc2pdf\csvem.csv")

for l in ls:
    temp = list(l.values())
    print(l)
    print(f"{temp[0]}_{ls.index(l)+1}")
