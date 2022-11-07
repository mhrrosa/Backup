import json
import random
from collections import OrderedDict
import re
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from xml.etree import ElementTree
import pandas
from unidecode import unidecode


def verifica(file_name):
    df = pandas.read_excel(file_name)

    for dados in df.iterrows():
        turma = str(dados[1]['Turma'])
        nome_disciplina = str(dados[1]['Nome da disciplina'])
        abrev_disciplina = str(dados[1]['Sigla da disciplina'])
        local_aula = str(dados[1]['Local de Aula'])
        nome_professor = str(dados[1]['Nome do Professor'])
        abrev_professor = str(dados[1]['Nome abreviado do professor'])


        if len(turma)> 30:
            print(turma,'Passou de 30 caracteres')
        if len(nome_disciplina) > 100:
            print(nome_disciplina,'Passou de 100 caracteres')
        if  len(abrev_disciplina) > 10:
            print(abrev_disciplina,'Passou de 10 caracteres')
        if  len(local_aula) > 50:
            print(local_aula,'Passou de 50 caracteres')
        if  len(nome_professor) > 100:
            print(nome_professor,'Passou de 100 caracteres')
        if  len(abrev_professor) > 10:
            print(abrev_professor,'Passou de 10 caracteres')



def main():
    # selecionando arquivos
    print('Selecione o arquivo turma disciplina')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])

    verifica(file_name_xlsx)

if __name__ == '__main__':
    f = open(f"Log.txt", "w+", encoding='utf8')
    main()
    f.close()
