
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
    class_id = requests.get(f'https://api.powercubus.com.br/v1/classes/list?timetable_id={project_id}',
                            headers=my_headers)

    return class_id.json()





def lista_Class(class_json, project_id, access_token):
    my_headers = {

        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }



    for classes in class_json:

            # comparando o name do xml prime com o name da class no powercubus
            # se for verdade verificar as id



            dict_valores = {

                    'grid_id': int(classes['grid_id']),
                    'unit_id': int(classes['unit_id']),
                    'name': str(classes['name']),
                    'grade': str(classes['grade']),
                    'course': str(classes['course']),
                    'external_id':  str(classes['external_id']),
                    'id': int(classes['id'])

            }
            valores_json = json.dumps(dict_valores)
            log_nameclass.write(str(classes['name'])+'\n')


            log_allclasses.write(valores_json+'\n')

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

    class_json = get_class(project_id, access_token)

    lista_Class(class_json, project_id, access_token)


if __name__ == "__main__":
    log_allclasses = open("log_AllClass.txt", "w+", encoding='utf-8')
    log_nameclass = open("log_NameClass.txt", "w+", encoding='utf-8')
    main()
    log_allclasses.close()
    log_nameclass.close()

