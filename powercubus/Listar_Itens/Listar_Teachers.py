
import json
from lxml import etree
import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import requests
from unicodedata import normalize


def remover_acentos(txt):
    return normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')



def get_class(project_id, access_token):
    my_headers = {'Authorization': f'Bearer {access_token}'}
    class_id = requests.get(f'https://api.powercubus.com.br/v1/teachers/list?timetable_id={project_id}',
                            headers=my_headers)

    return class_id.json()





def lista_Teacher(teacher_json, project_id, access_token):
    my_headers = {

        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }

    for teachers in teacher_json:

            # comparando o name do xml prime com o name da class no powercubus
            # se for verdade verificar as id



            dict_valores = {

                    'name': str(teachers['name']),
                    'external_id':  str(teachers['external_id']),
                    'id': int(teachers['id'])

            }
            valores_json = json.dumps(dict_valores)
            log_nameteachers.write(str(teachers['name'])+'\n')
            log_abrevteachers.write(str(teachers['abreviation'])+'\n')


            log_allteachers.write(valores_json+'\n')

def corrige_caracteres_especiais(file_name):
    # Lê o arquivo
    with open(file_name, "r", encoding="windows-1252") as file:
        file_data = file.read()

    # Substitui o símbolo que causa erro
    file_data = file_data.replace(" & ", " &amp; ")

    # Sobrescreve o arquivo
    with open(file_name, "w") as file:
        file.write(file_data)


def main():




    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    teachers_json = get_class(project_id, access_token)

    lista_Teacher(teachers_json, project_id, access_token)


if __name__ == "__main__":
    log_allteachers = open("log_AllTeachers.txt", "w+", encoding='utf-8')
    log_nameteachers = open("log_NameTeachers.txt", "w+", encoding='utf-8')
    log_abrevteachers = open("log_AbrevTeachers.txt", "w+", encoding='utf-8')
    main()
    log_allteachers.close()
    log_nameteachers.close()

