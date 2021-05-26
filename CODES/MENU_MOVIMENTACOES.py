import glob
import pandas as pd
import ATIVIDADES.MOVIMENTACOES as MOVIMENTACOES
import func_genericas.FORMATANDO_EXCEL as FORMATANDO_EXCEL
import func_genericas.CLASS_DIRETORIOS as CLASS_DIRETORIOS
import os
from datetime import date

def main(seguradora,razao_social,nome_fantasia,tipo_movimentacao,tabela,sistema,susep,estipulante,modalidade,endereco):
    
    # CHAMANDO CLASSE PARTE 1
    return_class = CLASS_DIRETORIOS.DIRETORIOS()

    # CHAMANDO CLASSE PARTE 2
    return_class.definindo_path_atual(seguradora,razao_social,nome_fantasia,tabela,sistema)
    path_atual = return_class.path_atual
    path_tabela = return_class.path_tabela
    path_protocolo = return_class.path_protocolo
    
    # MOSTRA DESTINO
    destino_base = '/'.join(path_atual.replace('\\','/').split('/')[-4:])
    print(f'\n\nO DESTINO DA BASE É: {destino_base}')
    
    # DEFINE DATA DE FECHAMENTO DA PREVISÃO PARA AVALIAÇÃO DE ADESÕES
    ano_mes_fechamento = path_atual.replace('\\','/').split('/')[-4:][2]
    fechamento_previsao = date(
                            year=int(ano_mes_fechamento.split('_')[1:][0]),
                            month=int(ano_mes_fechamento.split('_')[1:][1]),
                            day=25
    )
    
    # AVISA QUE DEVE SER INSERIDO NO BLACKBIRD
    if sistema == 'BLACKBIRD':
        print('\n\n\n\n       ##### REALIZAR NO SISTEMA BLACKBIRD #####\n\n')

    # CHAMANDO CLASSE PARTE 3
    return_class.definindo_labels(razao_social,tabela,susep,estipulante,modalidade,endereco,seguradora)
    label_1 = return_class.label_1
    label_2 = return_class.label_2
    label_3 = return_class.label_3

    # RODANDO FUNÇÕES
    if tipo_movimentacao == 4:
        protocolo_atual = max(glob.glob(path_protocolo + '*.xlsx'))
        df_atual = MOVIMENTACOES.main(path_atual,path_tabela,fechamento_previsao,tipo_movimentacao,sistema,seguradora,protocolo_atual)
        FORMATANDO_EXCEL.main(path_atual,path_atual,df_atual,label_1,label_2,label_3,'PRISMA')
    else:
        df_atual = MOVIMENTACOES.main(path_atual,path_tabela,fechamento_previsao,tipo_movimentacao,sistema,seguradora)
        FORMATANDO_EXCEL.main(path_atual,path_atual,df_atual,label_1,label_2,label_3)


    # SALVA A RAZÃO SOCIAL NO ARQUIVO DE MOVIMENTAÇÕES DO DIA
    movs_today(return_class.path_movimentacoes_do_dia,seguradora + '_' + razao_social)
    print('\n')


def perpetuo():
    
    return_class = CLASS_DIRETORIOS.DIRETORIOS()
    path_todos_storages = return_class.lista_storages
    
    perguntar_storage = 'S'
    while True:
        if perguntar_storage == 'S':
            # ARMAZENANDO VARIÁVEIS DO MENU
            seguradora,razao_social,nome_fantasia,tipo_movimentacao,tabela,sistema,susep,estipulante,modalidade,endereco = define_variaveis(path_todos_storages)
        else:
            tipo_movimentacao = define_movimentacao()
        
        main(seguradora,razao_social,nome_fantasia,tipo_movimentacao,tabela,sistema,susep,estipulante,modalidade,endereco)
        
        # PERGUNTANDO SE FEZ NO BLACKBIRD
        if sistema == 'BLACKBIRD':
            while True:
                relizado_blackbird = input('\n\nAs movimentações foram realizadas no blackbird/enviadas para seguradora em caso de endosso? (S ou N): ').upper()
                if relizado_blackbird not in ['S','N']:
                        print("   Digite apenas 'S' ou 'N'")
                elif relizado_blackbird == 'N':
                        print("   Faça as movimentações no Blackbird para poder continuar a movimentar")
                else:
                    break

        # PERGUNTANDO PROCESSO
        while True:
            try:
                continuar_processo = int(input(
            "\n****************************************\n\n\
        1 - CONTINUAR MOVIMENTAÇÕES COM O MESMO STORAGE\n\n\
        2 - CONTINUAR MOVIMENTAÇÕES COM OUTRO STORAGE\n\n\
        3 - PARAR MOVIMENTAÇÕES\n\n*****************************************\nSelecione a opção desejada: "))
                if continuar_processo not in range(1,4):
                    print('\n\n         Selecione uma das opções oferecidas!!')
                elif tipo_movimentacao == 4 and continuar_processo == 1:
                    print('\n\n         Para Storage Prisma, parar de movimentar ou selecionar outro storage!!')
                else:
                    break
            except ValueError:
                    print('\n\n         Digite um número!!')
        
        if continuar_processo == 1:
            perguntar_storage = 'N'
        elif continuar_processo == 2:
            perguntar_storage = 'S'
        elif continuar_processo == 3:
            break

def define_variaveis(path_todos_storages):

    # PEDINDO TIPO DE MOVIMENTAÇÃO
    tipo_movimentacao = define_movimentacao()

    # ESCOLHENDO SEGURADORA
    seguradora_escolhida_para_filtro = define_seguradora()

    # LENDO BASE DE STORAGES
    storages = pd.read_excel(path_todos_storages)
    storages.sort_values(by='NOME FANTASIA',inplace=True)

    # FILTRA OPÇÕES
    if seguradora_escolhida_para_filtro == 1:
        storages = storages.loc[storages['SEGURADORA']=='PORTO'].copy()
    elif seguradora_escolhida_para_filtro == 2:
        storages = storages.loc[storages['SEGURADORA']=='LIBERTY'].copy()

    # ANALISA SE TIPO DE MOVIMENTAÇÃO É PRISMA
    if tipo_movimentacao == 4:
        storages = storages.loc[storages['SISTEMA']=='PRISMA'].copy()
    else:
        storages = storages.loc[storages['SISTEMA']!='PRISMA'].copy()

    # CRIANDO LISTAS DE RAZÃO E FANTASIA
    susep = list(storages['SUSEP/C.O'].astype(str))
    tabela = list(storages['TABELA'].astype(str))
    sistema = list(storages['SISTEMA'])
    endereco = list(storages['ENDEREÇO'])
    modalidade= list(storages['MODALIDADE'].astype(str))
    seguradora = list(storages['SEGURADORA'])
    estipulante = list(storages['ESTIPULANTE'].astype(str))
    razao_social = list(storages['RAZÃO SOCIAL'])
    nome_fantasia = list(storages['NOME FANTASIA'])

    # PEDINDO STORAGE A SER MOVIMENTADO
    print('\n\n\n***************************************************************')
    for num,storage in enumerate(nome_fantasia):
        print('\t',str(num+1).zfill(2),'-', storage.upper())
    print('\n***************************************************************')

    while True:
        try:
            opcao = int(input('Digite o número do Storage que deseja realizar a movimentação: '))
            if opcao not in range(1,len(nome_fantasia)+1):
                print('\n\n         Selecione uma das opções oferecidas!!')
            else:
                opcao = opcao - 1
                break
        except ValueError:
                print('\n\n         Digite um número!!')


    susep = susep[opcao]
    tabela = tabela[opcao]
    sistema = sistema[opcao]
    endereco = endereco[opcao]
    seguradora = seguradora[opcao]
    apelido = nome_fantasia[opcao]
    modalidade = modalidade[opcao]
    estipulante = estipulante[opcao]
    storage_selecionado = razao_social[opcao]

    return seguradora,storage_selecionado,apelido,tipo_movimentacao,tabela,sistema,susep,estipulante,modalidade,endereco

def define_movimentacao():

    # PEDINDO TIPO DE MOVIMENTAÇÃO
    while True:
        try:
            tipo_movimentacao = int(input(
            "\n****************************************\n\n\
            1 - ADESÃO\n\n\
            2 - ALTERAÇÃO\n\n\
            3 - CANCELAMENTO\n\n\
            4 - PRISMA\n\n*****************************************\nDigite o tipo de movimentação que deseja realizar: "))
            if tipo_movimentacao not in range(1,5):
                print('         Selecione uma das opções oferecidas!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')

    return tipo_movimentacao

def define_seguradora():

    # PEDINDO SEGURADORA
    while True:
        try:
            escolha = int(input(
            "\n****************************************\n\n\
            1 - PORTO SEGURO\n\n\
            2 - LIBERTY\n\n\
            3 - AMBOS\n\n*****************************************\nDigite a seguradora que deseja: "))
            if escolha not in range(1,4):
                print('         Selecione uma das opções oferecidas!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')

    return escolha


def movs_today(path_mov_day,texto):

    if not os.path.exists(path_mov_day):
        df = pd.DataFrame(columns=['STORAGE'])
    else:
        df = pd.read_excel(path_mov_day)
    
    df_to_add = pd.DataFrame([texto.replace("\\",'/')],columns=['STORAGE'])
    df = df.append(df_to_add)
    df.drop_duplicates(subset=['STORAGE'],inplace=True)
    df.sort_values(by=['STORAGE'],inplace=True)
    df.to_excel(path_mov_day,index=False)

    print(f'{texto} ARMAZENADO EM MOVIMENTAÇÕES DO DIA')

if __name__ == "__main__":
    perpetuo()