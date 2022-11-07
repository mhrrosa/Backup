'''

1 - armazenar xml
2 - armazenar teachers.json(oq vem do powercubus via API)
3 - converter o name do xml para sem acentos e tudo maiusculo
4 - comparar o name do xml com o name do powercubus se for igual seguir para o proximo passo
 4.1 - se for igual atualizar o external_id do powercubus para o partner_id do xml

 Desenvolvido por Matheus Rosa
 Última atualização: 16/05/2022
'''
import json
import pandas
from lxml import etree
import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import requests
from unicodedata import normalize




def remover_acentos(txt):
    return normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')

def get_file_name():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('XML', '.xml')])
    return file_name


def get_teachers(project_id, access_token):

    my_headers = {'Authorization': f'Bearer {access_token}'}
    teachers_id = requests.get(f'https://api.powercubus.com.br/v1/teachers/list?timetable_id={project_id}',
                            headers=my_headers)

    return teachers_id.json()





def update_Abrev(file_name_xlsx, teachers_json, project_id, access_token):
    my_headers = {

        'Authorization': f'Bearer {access_token}',
        'Content-type': 'application/json',
        'Accept': 'application/json'
    }

    df = pandas.read_excel(file_name_xlsx)
    nomes_nao_encontrados =[]
    name_encontrado = []
    dict_retorna_abrev ={}
    for dados in df.iterrows():
        nome = dados[1]['Nome do Professor']
        abrev = dados[1]['Nome abreviado do professor']
        dict_retorna_abrev[nome] = abrev

    for teachers in teachers_json:

            try:
                abrev = dict_retorna_abrev[teachers['name']]
            except:
                f.write(teachers['name'])
                continue

            dict_valores ={
                                "abreviation": abrev,
                                "name": teachers['name'],
                                "maximum_daily_lessons":teachers['maximum_daily_lessons'],
                                "working_days": teachers['working_days'],
                                "idletimes": teachers['idletimes'],
                                "external_id": teachers['external_id'],
                                "id": teachers['id']
                        }
            valores_json = json.dumps(dict_valores)


            response = requests.put(f'https://api.powercubus.com.br/v1/teachers/update?'
                                    f'timetable_id={project_id}',
                                    headers=my_headers,data=valores_json)


            print(f'{teachers["name"]: <100} Atualização da ID retornou código '
                  f'{response.status_code} ({response.reason})')





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
    print('Selecione o arquivo turma disciplina')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])



    project_id = input('ID do Projeto: ')
    access_token = input('Token de acesso: ')

    teachers_json = get_teachers(project_id, access_token)

    update_Abrev(file_name_xlsx, teachers_json, project_id, access_token)


if __name__ == "__main__":
    f = open("log_Atualizaabrev.txt", "w+", encoding='utf-8')
    main()
    f.close()
    print('Atualização finalizada')
