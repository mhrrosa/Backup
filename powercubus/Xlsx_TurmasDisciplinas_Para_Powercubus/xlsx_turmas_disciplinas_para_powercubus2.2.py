""""
Script para realizar a conversão de um arquivo excel
contendo as turmas e as disciplinas e de um arquivo XML das turmas, disciplinas e salas
para o template utilizado no power cubus

Desenvolvido por Matheus Rosa e Fernando Dias
Última atualização: 21/10/2022
"""
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


def cria_abreviacao_professor(nome, dict_abreviacao):

    #criando a abreviaçao do professor com base no nome
    if nome == 'SEM PROFESSOR':
        return f'S. Prof ' \
               f'{random.choice(string.ascii_uppercase + string.digits)}' \
               f'{random.choice(string.ascii_uppercase + string.digits)}'
    if nome == 'nan':
        return f'S. Prof ' \
               f'{random.choice(string.ascii_uppercase + string.digits)}' \
               f'{random.choice(string.ascii_uppercase + string.digits)}'
    if nome in dict_abreviacao:
        return dict_abreviacao[nome]
    else:
        try:
            nome = nome.strip()
            abreviacao = nome.split(' ')[0][0:8] + ' ' + nome.split(' ')[-1][0]

            while abreviacao in dict_abreviacao.values():
                abreviacao = abreviacao[:-1] + random.choice(string.ascii_uppercase)

            dict_abreviacao[nome] = abreviacao
        except:
            print('erro no nome do professor: ', nome, '')
            abreviacao = ''
        return abreviacao


def cria_sigla_disciplina(nome, dicionario):
    # Converte numerais romanos no final do nome
    nome = re.sub(' I$', ' 1', nome)
    nome = re.sub(' II$', ' 2', nome)
    nome = re.sub(' III$', ' 3', nome)
    nome = re.sub(' IV$', ' 4', nome)
    nome = re.sub(' V$', ' 5', nome)
    nome = re.sub(' VI$', ' 6', nome)
    nome = re.sub(' VII$', ' 7', nome)
    nome = re.sub(' VIII$', ' 8', nome)

    # Converte todos os whitespaces para um espaço
    # Serve para remover múltiplos espaços seguidos

    nome = nome.replace('-', ' ')
    nome = ' '.join(nome.split())

    # Se o nome da disciplina já possui menos de 10 caracteres, utiliza o nome como sigla
    if len(nome) < 10:
        return nome

    # Remove palavras que começam com lowercase
    palavras = ' '.join(filter(lambda x: not x[0].islower(), nome.strip().split(' ')))
    palavras = palavras.split(' ')

    # Pega as iniciais de cada palavra e o máximo possível da primeira palavra sem ultrapassar o limite de 10 caracteres
    # Exemplo: 'Leitura e Escrita de Textos Técnico-Científicos' fica 'Leitu ETTC'
    sigla = ''
    for palavra in palavras[1:6]:
        sigla += palavra[0]
    sigla = palavras[0][0:9 - len(sigla)] + ' ' + sigla

    sigla = sigla.upper()

    while sigla in dicionario.values():
        sigla = sigla[:-1] + random.choice(string.ascii_uppercase)

    return sigla

#tratando o xml e transformando e lista para ser utilizado
def armazenaDadosXml(file_name_xml):
    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()

    #criando função para converter formato de horarios xml no formato powercubus
    def convertePeriod(period):

        if period == "1":
            valor = "7:05-7:50"
        elif period == "2":
            valor = "7:50-8:35"
        elif period == "3":
            valor = "8:35-9:20"
        elif period == "4":
            valor = "9:40-10:25"
        elif period == "5":
            valor = "10:25-11:10"
        elif period == "6":
            valor = "11:10-11:55"
        elif period == "7":
            valor = "11:55-12:40"
        elif period == "8":
            valor = "12:40-13:25"
        elif period == "9":
            valor = "13:25-14:10"
        elif period == "10":
            valor = "14:10-14:55"
        elif period == "11":
            valor = "15:15-16:00"
        elif period == "12":
            valor = "16:00-16:45"
        elif period == "13":
            valor = "16:45-17:30"
        elif period == "14":
            valor = "17:30-18:15"
        elif period == "15":
            valor = "18:15-19:00"
        elif period == "16":
            valor = "19:00-19:45"
        elif period == "17":
            valor = "19:45-20:30"
        elif period == "18":
            valor = "20:45-21:30"
        elif period == "19":
            valor = "21:30-22:15"
        elif period == "20":
            valor = "22:15-23:00"
        else:
            valor = ""

        return valor

    #criando funççao para converter formato dia xml para formato powercubus
    def convertDays(day):
        if day == "1111111" or day == "111111":
            valor = "Cada dia"

        elif day == "1000000" or day == "100000":
            valor = "Segunda-feira"

        elif day == "0100000" or day == "010000":
            valor = "Terça-feira"

        elif day == "0010000" or day == "001000":
            valor = "Quarta-feira"

        elif day == "0001000" or day == "000100":
            valor = "Quinta-feira"

        elif day == "0000100" or day == "000010":
            valor = "Sexta-feira"

        elif day == "0000010" or day == "000001":
            valor = "Sábado"

        elif day == "0000001":
            valor = "Horario Flutuante"
        else:
            valor = ""

        return valor

    #armazenando id e nome da disciplina
    dic_subject = {}
    for subject in root.iter("subject"):
        id = subject.get("id")
        name = subject.get("name")

        dic_subject[id] = name
    #armazenando id e nome dos professores
    dic_teacher={}
    for teacher in root.iter("teacher"):
        id = teacher.get("id")
        name = teacher.get("name")

        dic_teacher[id] = name
    #armazenando id, nome e abreviação da turma
    dic_class = {}
    for classes in root.iter("class"):
        id = classes.get("id")
        name = classes.get("name")
        short = classes.get("short")

        dic_class[id] = name, short

    #armazenando id e nome dos grupos
    dict_retorna_group = {}
    for groups in root.iter("group"):
        id = groups.get('id')
        name = groups.get('name')

        dict_retorna_group[id] = name
    #armazenando id,nome e abreviações da sala
    dict_classrooms = {}
    for classrooms in root.iter("classroom"):
        id = classrooms.get("id")
        name = classrooms.get("name")
        short = classrooms.get("short")

        dict_classrooms[id] = name,short

    #armazenando id da aula, periodo e dia dos cartões, convertendo periodo e dia para o formato powercubus
    dict_cards = {}
    for cards in root.iter("card"):
        # parametros
        id = str(cards.get('lessonid'))
        period = cards.get('period')
        day = cards.get('days')

        #convertento o periodo e dia
        valor = f"({convertePeriod(period)}/{convertDays(day)})"
        try:
            valor_dict = dict_cards[id]

            if valor_dict.__contains__(valor):
                continue
            else:
                dict_cards[id] += valor
        except:

            dict_cards[id] = valor

    #criando a lista final do xml
    valores = []
    for lessons in root.iter("lesson"):
        #armazenando ids
        id = lessons.get('id')
        subjectid = lessons.get('subjectid')
        groupid = lessons.get('groupids')
        classid = lessons.get('classids')
        classroomsid = lessons.get('classroomids')
        teacherid = lessons.get('teacherids')

        #verificando a quantidade de professores
        if teacherid.count(',') > 0:
            teacherid = teacherid.split(",")[0]

        #armazenando a nome da disciplina
        name_subject = dic_subject[subjectid]

        #armazenando o cr do curso
        cr_curso = name_subject.split(";")[0]

        #verificando se o professor existe na aula, se tiver armazenando o nome
        try:
            teacher = dic_teacher[teacherid]
        except:
            print(teacherid, 'sem professor')
            teacher = "sem professor"

        #verificando a quantidade de salas, se tiver mais que uma são separadas uma para cadda aula
        quantidade_classroom = classroomsid.count(',')
        # verificando a quantidade de turmas, se tiver mais que uma são separadas uma para cadda aula
        quantidade_classid = classid.count(',')
        #iniciando o valor da posicao dos itens
        posicao_sala = 0
        posicao_turma = 0

        try:
            #criando um laço com a quantidade de turmas
            for i in range(0, quantidade_classid + 1):
                # criando um laço com a quantidade de salas
                for x in range(0, quantidade_classroom + 1):

                    #se não existir sala a aula é ignorada
                    if classroomsid == "":
                        continue
                    #armazena os valores da sala
                    sala = str(dict_classrooms[classroomsid.split(',')[posicao_sala]]).split(',')[0]
                    abrev_sala = str(dict_classrooms[classroomsid.split(',')[posicao_sala]]).split(',')[1]
                    #armazena os valores da turma
                    turma = str(dic_class[classid.split(',')[posicao_turma]]).replace(',', '/')
                    periodo = turma.split("/")[0].split(";")[1]

                    #armazenando o dia e hora da aula
                    try:
                        dia_hora = dict_cards[id]
                    except KeyError:
                        f.write(f"id {id} não teve card encontrado")
                        continue
                    #armazenando o grupo
                    try:
                        group = dict_retorna_group[groupid.split(',')[0]]
                    except:
                        f.write('erro no group da lesson' + id + '\n')
                        continue

                    #convertendo o formato entire class e turma completa para teorico
                    if group.replace(' ', '').upper() == 'TURMACOMPLETA' or group.replace(' ','').upper() == 'ENTIRECLASS':
                        group = 'Teórico'

                    #aramazenando a quantidade de dias e hora para separar para cada aula
                    quantidade_dia = dia_hora.count("/")

                    try:
                        for posicao_sala in range(0, quantidade_dia):
                            #armazenando o dia na posição correta
                            dia = dia_hora.split(')')[posicao_sala].split('/')[1]
                            #convertendo o turno da turma
                            turma = turma.replace('Manhã e Tarde', 'I')
                            #criando os valores da lista
                            valor = f"{cr_curso}|{periodo}|{turma}|{sala}|{abrev_sala}|{dia_hora}|{name_subject}|{dia}|{group}|{teacher}"
                            #armazenando o valor na lista
                            valores.append(valor)
                    except:
                        continue
                posicao_sala += 1
                posicao_turma += 1
        except:
            print('erro ',classroomsid)
    return valores


def gerarRelatorio(file_name_xlsx, file_name_xml, file_json_sigla,file_json_turmasfic):
    #armazenando os arquivos auxiliares em dict
    siglas = dict(file_json_sigla)
    turmas_fic = dict(file_json_turmasfic)
    #convertendo o nome da turma em siglas
    def converteSigla(turma):
        if turma in siglas.keys():
            return siglas[turma]
        else:
            return turma

    # função para converter o turno em letras
    def converte_turno(turno):

        if turno == "MANHÃ":
            turno = "M"
        elif turno == "NOITE":
            turno = "N"
        elif turno == "MANHÃ E TARDE":
            turno = 'I'
        elif turno == "INTEGRAL":
            turno = 'I'

        return turno

    # função para calcular a modulação
    def calcula_modulacao(nome_turma, modulacao_matriz):
        # tratamento da modulação da matriz
        '''
            Modulação padrão = 60
            Modulação Medicina = 90

            Exemplo
            modulação padrão 60 = 1
            modulação padrão 30 = 2

        '''
        # criando a modulção padrão
        numero_divisao = 60
        # se for medicina e tratado como 90 a modulação
        if nome_turma.__contains__('MED'):
            numero_divisao = 90
            modulacao = int(numero_divisao / modulacao_matriz)

        else:
            modulacao = int(numero_divisao / modulacao_matriz)
        return modulacao

    #função para converter o tipo atividade para o formato modulação
    def converte_tipo_atividade(tipo_atividade):
        if tipo_atividade == 'Prático':
            tipo_atividade = 'P'
        elif tipo_atividade == 'Tutoria':
            tipo_atividade = 'Tut'
        else:
            tipo_atividade = 'Teórico'

        return tipo_atividade

    def gera_modulacao(tipo_atividade,modulacao_matriz,modulacao_curso):
        if tipo_atividade == 'Teórico':
            cod_origem_aula = f"{tipo_atividade}"
        else:
            cod_origem_aula = f"{tipo_atividade}{modulacao_matriz}-{modulacao_curso}"

        return cod_origem_aula

    #criando uma sala generica caso não seja encontrada
    def valida_sala(num,key,dict):

        try:
            sala_xml = dict[key]
            return sala_xml
        except:
            # se não encontrar deve-se criar uma sala generica SEM SALA 01
            num += 1
            sala_xml = f'Sem Sala {num}'
            f.write('Sala nao foi encontrada:' + key + '\n')
            return sala_xml

    # Se o nome do professor conter caracteres não aceitaveis, cria-se um professor generico SEM PROFESSOR 02
    def valida_professor(num,nome_professor):
        #lista de erros encontrados que o conversor deve ignorar
        lista_erros =[
                      "/",
                      "(",
                      ",",
                      "SEM PROFESSOR",
                      "Sem Professor"
                      ]

        for erro in lista_erros:
            if nome_professor.__contains__(erro):
                num += 1
                nome_professor = f"Sem Professor {nome_professor}"
                return nome_professor
            else:
                return nome_professor

    def valida_professor_abrev(num,key,dict):
        if key.__contains__("Sem Professor"):
            abrev_professor = f"SP {num}"
        else:
            abrev_professor = dict[key]
        return abrev_professor
    
    # transformando o caminho em um read de excel
    df = pandas.read_excel(file_name_xlsx)
    lista_xml = armazenaDadosXml(file_name_xml)



    # tratando os dados nulos(nan) do excel
    df['Docente 2022.2'].fillna('Sem Professor', inplace=True)
    df['Nome da disciplina'].fillna('Sem Disciplinas', inplace=True)
    df['Turma'].fillna('Sem Turma', inplace=True)
    df['Cód. do Curso'].fillna('Sem Turma', inplace=True)
    df['Modulação pela Matriz'].fillna('1', inplace=True)
    df['Turno'].fillna('Sem turno', inplace=True)


    # dicts para armazenamento
    dic_disciplina = {}
    dic_turma = {}
    dic_grupo = {}
    dic_sala = {}
    dict_siglas = {}
    dict_abreviacao = {}
    dict_modulacoes = {}
    dict_agrupamento = {}
    dict_retorna_maior_div= {}
    dict_retorna_professor= {}
    dict_retorna_salaxml= {}

    #iniciando o valor inicial do agrupamento
    agrupamento = 0

    #se não houver os itens, deve ser criado um nome generico, utilizamos essa variavel para a criação de um nome generico
    num_professor = 0
    num_sala = 0

    #criando o valor inicial da chave do dict
    key_final = 0
    #dict que vai ser transformado na planilha
    dict_dados_final = {}

    #variaveis para retirar a duplicação do dict
    dit_sem_duplicadas ={}
    repetidos =[]

    #print para monstrar no console em qual parte esta a execução
    print('montando dict excel...')

    for dados in df.iterrows():

        #criando variaveis e adicionando os valores
        periodo = dados[1]['Período']
        periodo = str(periodo).split('.')[0]
        turma = dados[1]['Turma']
        turno = dados[1]['Turno']
        nome_disciplina = dados[1]['Nome da disciplina'].upper()
        nome_disciplina_corrigido = nome_disciplina.replace(" ", "").replace('-', "")
        nome_disciplina_comparacao = unidecode(nome_disciplina_corrigido)
        tipo_atividade = dados[1]['Tipo de Atividade']
        nome_curso = dados[1]['Nome do Curso']
        nome_professor = dados[1]['Docente 2022.2'].upper()
        modulacao_curso = str(dados[1]['Modulação']).split('.')[0]

        # tratamento de dados para o formatado do cód de origem da aula
        try:
            modulacao_matriz = int(dados[1]['Modulação pela Matriz'])
        except ValueError:
            print('1 - erro para encontrar a modulação: ', nome_curso)
            continue



        # tuma completa não possui volores para modulação matriz e modulação curso
        # criando siglas professores
        if nome_professor not in dict_abreviacao:
            dict_abreviacao[nome_professor] = cria_abreviacao_professor(nome_professor, dict_abreviacao)

        # criando siglas disciplinas
        if nome_disciplina not in dict_siglas:
            dict_siglas[nome_disciplina.strip()] = cria_sigla_disciplina(nome_disciplina, dict_siglas)

        #convertendo o nome da turma para sigla e armazenando o curso
        nome_turma = converteSigla(nome_curso)
        turno = turno.split(' ')[0].upper()
        # converte o turno
        turno = converte_turno(turno)

        #criando a variavel que vai ser utilizado na chave do dict
        turma_comparacao = f"{nome_turma} - {periodo}{turma} - {turno}"

        # tratamento da modulação da matriz
        modulacao_matriz = calcula_modulacao(nome_turma,modulacao_matriz)

        #convertendo o tipo de atividade para ficar no formato modulação
        tipo_atividade = converte_tipo_atividade(tipo_atividade)

        #criando a modulação que vai ser armazenanada na cod de origem no powercubus
        cod_origem_aula = gera_modulacao(tipo_atividade,modulacao_matriz,modulacao_curso)

        if tipo_atividade == 'Teórico':
            modulacao_matriz = 'Teórico'

        #criando dictr para armazenaar todas as modulações da turma
        try:
                dict_modulacoes[turma_comparacao] += f"|{modulacao_matriz}"
        except:
                dict_modulacoes[turma_comparacao] = f"{modulacao_matriz}"

        #criando a chave do dict agrupamento e armazenando no dict agrupamento
        key_agrupamento = f"{turma_comparacao}{nome_disciplina_comparacao}{tipo_atividade}{modulacao_matriz}"
        if cod_origem_aula == 'Teórico' or cod_origem_aula == 'P1-1':
            agrupamento += 1
            dict_agrupamento[key_agrupamento] = agrupamento

        #armazenanado a maior divisão, para multiplicaar a aula teorica e p1-1 ate essa quantidade
        if modulacao_matriz != 'Teórico':
            try:
                    num_mod = int(dict_retorna_maior_div[turma_comparacao])
                    if modulacao_matriz > num_mod:
                        dict_retorna_maior_div[turma_comparacao] = modulacao_matriz
            except:
                dict_retorna_maior_div[turma_comparacao] = modulacao_matriz

        #criando chave dos professores e armazenando no dict professores
        key_professor = f"{turma_comparacao}{nome_disciplina_comparacao}{modulacao_matriz}"
        try:
                dict_retorna_professor[key_professor] += f"|{nome_professor}"
        except:
                dict_retorna_professor[key_professor] = f"{nome_professor}"

    # organuzando os dados
    print('montando dict xml...')
    for x in lista_xml:
        # corrigir nome turma
        barra = r"'\'"
        turma = x.split('|')[2].split('/')[0].replace(barra, '').replace(barra, '').replace(barra, '').replace(barra,'').replace("'", '').replace('(', '')
        short = x.split('|')[2].split('/')[1].replace(barra, '').replace(barra, '').replace(barra, '').replace(barra,'').replace("'", '').replace('(', '')
        sala_xml = x.split("|")[3].replace('(', '').replace(";", "-").replace("'", '')

        #criando variaveis
        nome_turma = converteSigla(turma.split(";")[0])
        periodo_turma = turma.split(";")[1]
        tipo_turma = turma.split(";")[2]
        turno = turma.split(";")[3]
        turno = turno.split(' ')[0].upper()
        grupo_turma = x.split('|')[8]
        nome_professor = x.split("|")[9]

        #tratamento especial para turmas fic
        if turma.__contains__('Engenharia;') or turma.__contains__('Multicom'):
            #criando variaveis no formato turma fic
            cr_curso = short.split(";")[5]
            periodo = short.split(";")[2]
            tipo = short.split(";")[3]
            turno = short.split(";")[4]

            chave_turma_fic = f"{cr_curso}|{periodo}"

            if chave_turma_fic in turmas_fic.keys():
                name =  turmas_fic[chave_turma_fic]

                new_name = f"{name};{periodo};{tipo};{turno}"
                name = new_name
            else:
                f.write(f'Aviso[Info]: erro ao converter a multicom name ={turma} , short={short}''\n')
                continue

        #converte o turno
        turno = converte_turno(turno)
        #criando nome da turma convertido
        nome_turma_corrigido = f"{nome_turma} - {periodo_turma}{tipo_turma} - {turno}"

        # corrigir nome disciplina
        try:
            nome_disciplina = x.split('|')[6].split(';')[2]
            nome_disciplina_corrigido = nome_disciplina.replace(' ', '').upper().replace('-', '')
            nome_disciplina_corrigido = unidecode(nome_disciplina_corrigido)
        except:
            continue

        #criando chave para retornar a sala
        key_dados_xml = f"{nome_turma_corrigido}{nome_disciplina_corrigido}{grupo_turma}{nome_professor.split(';')[0]}"
        #armazenando no dict sala
        dict_retorna_salaxml[key_dados_xml] = sala_xml

    #print para mosntrar no log qual parte do codigo esta em execução
    print('montando dict final...')
    for dados in df.iterrows():
        # tratamento por cota do formato da planilha que as linhas continuam sem os dados
        cod_curso = str(dados[1]['Cód. do Curso'])
        #se não tiver turma a aula é ignorada
        if cod_curso == "Sem Turma":
            continue

       #criando as variaveis
        periodo = dados[1]['Período']
        periodo = str(periodo).split('.')[0]
        turma = dados[1]['Turma']
        turno = dados[1]['Turno']
        nome_disciplina = dados[1]['Nome da disciplina'].upper()
        nome_disciplina_corrigido = nome_disciplina.replace(" ", "").replace('-', "")
        nome_disciplina_comparacao = unidecode(nome_disciplina_corrigido)
        nome_curso = dados[1]['Nome do Curso']
        nome_professor = dados[1]['Docente 2022.2'].upper()
        carga_horaria = dados[1]['Carga Horária']
        tipo_atividade = dados[1]['Tipo de Atividade']
        ano_turma = '2022/2'
        # tratamento turno
        turno = turno.split(' ')[0].upper()
        # converte o turno
        turno = converte_turno(turno)
        # tratando dados
        nome_turma = converteSigla(nome_curso)
        # Conversão da modulação_curso para string
        modulacao_curso = str(dados[1]['Modulação']).split('.')[0]
        # tratamento de dados para o formatado do cód de origem da aula
        try:
            modulacao_matriz = int(dados[1]['Modulação pela Matriz'])
        except ValueError:
            print('2 - erro para encontrar a modulação: ', nome_curso)
            continue

        # tratamento da modulação da matriz, modulação 60 = 1 e modulação 30 = 2
        modulacao_matriz = calcula_modulacao(nome_turma,modulacao_matriz)

        #convertendo o tipo de atividade para ficar no formato modulação
        tipo_atividade = converte_tipo_atividade(tipo_atividade)

        # gerando o cód. de origem de origem da aula
        cod_origem_aula = gera_modulacao(tipo_atividade,modulacao_matriz,modulacao_curso)

        #criando variaveis que vão ser utilizadas  na chave
        turma_comparacao = f"{nome_turma} - {periodo}{turma} - {turno}"
        turma_corrigida = f"{nome_turma} - {periodo}{turma} - {turno} - {ano_turma}"

        #formatando o teorico para vazio
        if cod_origem_aula == 'Teórico':
            modulacao_matriz = ''

        '''
            Aulas teoricas e P1-1 da mesma turma vão ser compartilhadas entre si, com o mesmo cod de agrupamento
        '''
        #verificando a modulação
        if modulacao_matriz == '' or cod_origem_aula == 'P1-1':

            #armazenando a maior quantidade de modulações
            try:
                quantidade_modulacao = dict_retorna_maior_div[turma_comparacao]
            except:
                print(turma_comparacao,' so teorica')
                quantidade_modulacao = 1

            #repetindo as aulas para todas as quantidades de modulação
            for posicao in range(1, quantidade_modulacao+1):

                    #verificando se é teorico para criar a chave dos professores
                    if cod_origem_aula == 'Teórico':
                        key_professor = f"{turma_comparacao}{nome_disciplina_comparacao}{cod_origem_aula}"
                    else:
                        key_professor = f"{turma_comparacao}{nome_disciplina_comparacao}{posicao}"

                    #armazenando a lista de professores
                    try:
                        lista_professores = dict_retorna_professor[key_professor]
                    except:
                        lista_professores = f"Sem Professor {num_professor}"

                    #armazenanado a quantidade de professores
                    quantidade_professores = lista_professores.count('|')
                    if quantidade_professores == 0:
                        quantidade_professores = 1
                    else:
                        quantidade_professores += 1

                    for posicao_professor in range(0, quantidade_professores):
                        #nome do professor
                        nome_professor = lista_professores.split("|")[posicao_professor]

                        num_professor+=1
                        #Se o nome do professor conter caracteres não aceitaveis, cria-se um professor generico SEM PROFESSOR 02
                        nome_professor = valida_professor(num_professor,nome_professor)
                        #criando e validando abreviação
                        abrev_professor = valida_professor_abrev(num_professor,nome_professor,dict_abreviacao)


                        # definindo cod agrupamento
                        key_agrupamento = f"{turma_comparacao}{nome_disciplina_comparacao}{tipo_atividade}{posicao}"

                        #definindo chave da sala
                        key_sala = F"{turma_comparacao}{nome_disciplina_comparacao}{cod_origem_aula}{nome_professor}"
                        #cria uma sala generica caso não seja encontrada
                        num_sala += 1
                        sala_xml = valida_sala(num_sala, key_sala, dict_retorna_salaxml)

                        if tipo_atividade == 'Teórico':
                            #criando a chave
                            key_agrupamento = f"{turma_comparacao}{nome_disciplina_comparacao}{tipo_atividade}{tipo_atividade}"
                            cod_origem_aula = f"{tipo_atividade}"
                            #buscando o numero do agrupamento
                            agrupamento = dict_agrupamento[key_agrupamento]

                            # se for 1, significa que é turma base
                            '''
                                Turma Base =  Turma pagante
                            '''
                            if posicao == 1:
                                cod_origem_aula += ' Base'
                            else:
                                cod_origem_aula = "Teórico"
                        elif tipo_atividade=="P" and posicao == 1:
                            cod_origem_aula = f"{tipo_atividade}{posicao}-{modulacao_curso}"
                            agrupamento = dict_agrupamento[key_agrupamento]
                        if quantidade_modulacao == 1:
                            agrupamento = ''
                        #criando dict que vai se tornar o excel final
                        key_final += 1
                        dict_dados_final[key_final] = {
                            'Turma': f'{turma_corrigida.replace("2022/2", "").replace("Turma Ficticia","Turma Fic").replace("INTERCAMBI","INTER")} div {posicao}',
                            'Etapa': dados[1]["Período"],
                            'Curso': nome_turma,
                            'Nome da disciplina': str(dados[1]["Nome da disciplina"]).strip(),
                            'Sigla da disciplina': dict_siglas[nome_disciplina.strip()],
                            'Area': '-',
                            'Nome do Professor': nome_professor,
                            'Nome abreviado do professor': abrev_professor,
                            'Email do Professor': "",
                            'Local de Aula': sala_xml,
                            'Aulas': carga_horaria,
                            'Máximo de aulas diárias': '',
                            'Agrupamento de aulas': '1',
                            'Máximo de dias de aula na semana': '',
                            'Permitir aulas em dias consecutivos': '1',
                            'Código de origem da turma': '',
                            'Código de origem da disciplina':'',
                            'Código de origem do professor': '',
                            'Código de origem do local': '',
                            'Código de origem da área': '',
                            'Nome da unidade da turma': 'Curitiba',
                            'Grupo de períodos': 'Graduação Presencial',
                            'Nome da unidade do local': '',
                            'Sigla da unidade da turma': 'Curitiba',
                            'Sigla da unidade do local': 'Curitiba',
                            'Código de origem da aula': cod_origem_aula,
                            'Código do agrupamento': agrupamento,
                            'Horário Fixo': '',
                        }
                        key_final += 1

        else:
            num_professor += 1
            # Se o nome do professor conter caracteres não aceitaveis, cria-se um professor generico SEM PROFESSOR 02
            nome_professor = valida_professor(num_professor, nome_professor)
            # criando e validando abreviação
            abrev_professor = valida_professor_abrev(num_professor, nome_professor, dict_abreviacao)

            # definindo chave da sala
            key_sala = F"{turma_comparacao}{nome_disciplina_comparacao}{cod_origem_aula}{nome_professor}"
            # cria uma sala generica caso não seja encontrada
            num_sala+=1
            sala_xml = valida_sala(num_sala, key_sala, dict_retorna_salaxml)

            # criando dict que vai se tornar o excel final
            key_final += 1
            dict_dados_final[key_final] = {
                'Turma': f'{turma_corrigida.replace("2022/2", "").replace("Turma Ficticia","Turma Fic").replace("INTERCAMBI","INTER")} div {modulacao_matriz}',
                'Etapa': dados[1]["Período"],
                'Curso': nome_turma,
                'Nome da disciplina': str(dados[1]["Nome da disciplina"]).strip(),
                'Sigla da disciplina': dict_siglas[nome_disciplina.strip()],
                'Area': '-',
                'Nome do Professor': nome_professor,
                'Nome abreviado do professor': abrev_professor,
                'Email do Professor': "",
                'Local de Aula': sala_xml,
                'Aulas': carga_horaria,
                'Máximo de aulas diárias': '',
                'Agrupamento de aulas': '1',
                'Máximo de dias de aula na semana': '',
                'Permitir aulas em dias consecutivos': '1',
                'Código de origem da turma': '',
                'Código de origem da disciplina': '',
                'Código de origem do professor': '',
                'Código de origem do local': '',
                'Código de origem da área': '',
                'Nome da unidade da turma': 'Curitiba',
                'Grupo de períodos': 'Graduação Presencial',
                'Nome da unidade do local': '',
                'Sigla da unidade da turma': 'Curitiba',
                'Sigla da unidade do local': 'Curitiba',
                'Código de origem da aula': cod_origem_aula,
                'Código do agrupamento': '',
                'Horário Fixo': '',
            }

    #retirando os repetidos
    for chave, valor in dict_dados_final.items():
        if valor not in repetidos:
            dit_sem_duplicadas[chave]= valor
            repetidos.append(valor)
        else:
            continue
    new_df = pandas.DataFrame(data=dit_sem_duplicadas)
    new_df = (new_df.T)
    new_df.to_excel(file_name_xlsx.replace('.xlsx', ' PowerCubus.xlsx'), sheet_name='Template', index=False)
def main():
    # selecionando arquivos
    print('Selecione o arquivo turma disciplina')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])

    print('Selecione o arquivo xml')
    Tk().withdraw()
    file_name_xml = askopenfilename(filetypes=[('xml', ".xml")])

    with open('../Arquivos_Auxiliares/siglas.json', 'r', encoding="utf-8") as j:
        file_json_sigla = json.loads(j.read())

    with open('../Arquivos_Auxiliares/turmasfic.json', 'r', encoding="utf-8") as j:
        file_json_turmasfic = json.loads(j.read())

    gerarRelatorio(file_name_xlsx, file_name_xml, file_json_sigla,file_json_turmasfic)


if __name__ == '__main__':
    f = open(f"Log.txt", "w+", encoding='utf8')
    main()
    f.close()
