from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas
import numpy
import xml.etree.ElementTree as ElementTree
import json

def lista_Names(file):
    tree_xml = ElementTree.parse(file)
    root = tree_xml.getroot()


    for teacher in root.iter("teacher"):
        name = teacher.get('name').split(';')[0]
        f.write(name+'\n')

def main():
    print('Selecione o arquivo XML ASC')
    Tk().withdraw()
    file_name_asc_xml = askopenfilename(filetypes=[('xml', '.xml')])

    lista_Names(file_name_asc_xml)


if __name__ == '__main__':
    f = open(f"Log_TeacherXml.txt", "w+", encoding='utf8')
    main()
    f.close()