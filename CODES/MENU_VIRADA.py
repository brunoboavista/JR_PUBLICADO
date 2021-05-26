import os
import time
import pandas as pd
import VIRADA.VIRAR_PLANILHAS as VIRAR_PLANILHAS
import VIRADA.FECHAR_FATURA as FECHAR_FATURA
import func_genericas.FORMATANDO_EXCEL as FORMATANDO_EXCEL
import func_genericas.CLASS_DIRETORIOS as CLASS_DIRETORIOS
from datetime import date
from dateutil.relativedelta import relativedelta

def main():

    # CHAMANDO MENUS
    ano,mes = menu()
    tipo_virada = menu_2()

    # DEFININDO VARIAVEIS DE DATA
    ano_mes_fechamento = date(ano,mes,1).strftime('P_%Y_%m')
    ano_mes_virada = (date(ano,mes,1) + relativedelta(months=1)).strftime('P_%Y_%m')
    fechamento_virada = date(
                            year=int(ano_mes_virada.split('_')[1:][0]),
                            month=int(ano_mes_virada.split('_')[1:][1]),
                            day=25
    )
    fechamento_fechamento = date(
                            year=int(ano_mes_fechamento.split('_')[1:][0]),
                            month=int(ano_mes_fechamento.split('_')[1:][1]),
                            day=25
    )

    # CHAMANDO CLASSE PARTE 1
    return_class = CLASS_DIRETORIOS.DIRETORIOS()
    path_todos_storages = return_class.lista_storages

    # ARMAZENANDO VARIÁVEIS DO MENU
    storages = pd.read_excel(path_todos_storages)
    storages.sort_values(by='NOME FANTASIA',inplace=True)

    # DAR UM LOC NOS CRITÉRIOS DE SEGURADORA, SISTEMA
    if tipo_virada == 1:
        storages = storages.loc[storages['SISTEMA']=='BLACKBIRD']
        clientes = 'BLACKBIRD'
    elif tipo_virada == 2:
        storages = storages.loc[storages['SISTEMA']=='PLANILHA']
        clientes = 'PLANILHA'
    elif tipo_virada == 3:
        storages = storages.loc[storages['SEGURADORA']=='PORTO']
        storages = storages.loc[storages['SISTEMA']=='PRISMA']
        clientes = 'PRISMA PORTO'
    elif tipo_virada == 4:
        storages = storages.loc[storages['SEGURADORA']=='LIBERTY']
        clientes = 'LIBERTY'
    elif tipo_virada == 99:
        storages = storages.loc[storages['SEGURADORA']=='PORTO']
        clientes = 'PORTO'

    
    # CRIANDO LISTAS DE RAZÃO E FANTASIA  
    susep = list(storages['SUSEP/C.O'].astype(str))
    tabela = list(storages['TABELA'].astype(str))
    # sistema = list(storages['SISTEMA'])
    endereco = list(storages['ENDEREÇO'])
    seguradora = list(storages['SEGURADORA'])
    modalidade = list(storages['MODALIDADE'].astype(str))
    estipulante = list(storages['ESTIPULANTE'].astype(str))
    razao_social = list(storages['RAZÃO SOCIAL'])
    nome_fantasia = list(storages['NOME FANTASIA'])

    # CONFIRMANDO DADOS PROCESSO
    print(f"\n\n\nINICIANDO VIRADA DOS CLIENTES {clientes}...")
    print("FECHANDO FATURAS COM DATA DE CORTE " + '25' + '/' + ano_mes_fechamento.split('_')[2] + '/' + ano_mes_fechamento.split('_')[1])
    print("CRIANDO PREVISÕES COM DATA DE CORTE " + '25' + '/' + ano_mes_virada.split('_')[2] + '/' + ano_mes_virada.split('_')[1])
    while True:
        print(f'\n\nO PROCESSO INICIARÁ OS FECHAMENTOS COM OS DADOS ACIMA')
        last_chance = input('\n     PARA CONFIRMAR APERTE "ENTER", PARA CANCELAR APERTE CTRL+C')
        if last_chance == '':
            break
    print('\n\n\n***************************************************')
    time.sleep(10)
    

    # INICIANDO PROCESSO
    for i in range(1,len(nome_fantasia)+1):
        susep_selected = susep[i-1]
        tabela_selected = tabela[i-1]
        # sistema_selected = sistema[i-1]
        endereco_selected = endereco[i-1]
        modalidade_selected = modalidade[i-1]
        seguradora_selected = seguradora[i-1]
        estipulante_selected = estipulante[i-1]
        razao_social_selected = razao_social[i-1]
        nome_fantasia_selected = nome_fantasia[i-1]

        # CHAMANDO CLASSE PARTE 2
        return_class.definindo_virada(seguradora_selected,
                                        razao_social_selected,
                                        nome_fantasia_selected,
                                        ano_mes_fechamento,
                                        ano_mes_virada)
        path_virada = return_class.path_virada
        dir_virada = return_class.dir_virada
        path_atual = return_class.path_atual
        dir_atual = return_class.dir_fechamento

        # CHAMANDO CLASSE PARTE 4
        return_class.definindo_labels(razao_social_selected,
                                    tabela_selected,
                                    susep_selected,
                                    estipulante_selected,
                                    modalidade_selected,
                                    endereco_selected,
                                    seguradora_selected)
        label_1 = return_class.label_1
        label_2 = return_class.label_2
        label_3 = return_class.label_3

        # PRINTA NOME DO STORAGE
        print(f'REALIZANDO VIRADA: {nome_fantasia_selected}')

        # VALIDA SE EXISTE O MÊS BASE A SER FECHADO
        if [os.path.exists(path_atual), os.path.exists(path_atual.replace('P_','F_'))] == [False,False]:
            print('O storage não tem base do período escolhido ou o nome está errado na planilha')    
        else:
            # VALIDA SE JÁ EXISTE O ARQUIVO VIRADO, SEJA COMO PREVISÃO OU FATURA
            if [os.path.exists(path_virada), os.path.exists(path_virada.replace('P_','F_'))] == [False,False]:
                continuar = 'S'
            else:
                continuar = 'N'
                print('Esta virada já foi realizada e não pode ser refeita pelo risco de perder a vigência de um cliente.')
               
            # RODA PROCESSO
            if continuar == 'S':

                # ARRUMA O PATH ATUAL
                if not os.path.exists(path_atual):
                    path_atual = path_atual.replace('P_','F_')

                # SEPARA PROCESSOS DEPENDENDO DA SEGURADORA/SISTEMA
                if seguradora_selected == 'LIBERTY':
                    df_virada = try_to_read_a_path(path_atual)
                else:
                    df_virada = VIRAR_PLANILHAS.main(path_atual,path_virada,dir_virada,fechamento_virada)
                
                # FORMATA EXCEL E PRINTA NA TELA
                FORMATANDO_EXCEL.main(path_virada,path_atual,df_virada,label_1,label_2,label_3,tipo='VIRADA')
                print("Virada realizada...")

                 # FECHANDO A FATURA
                if seguradora_selected != 'LIBERTY':
                    df_atual = FECHAR_FATURA.main(path_atual,fechamento_fechamento)
                    FORMATANDO_EXCEL.main(path_atual,path_atual,df_atual,label_1,label_2,label_3)
                    print("Fatura fechada...")

            else:
                print("Virada não realizada...")

            # RENOMEANDO DIRETÓRIO ANTIGO PARA FECHAMENTO
            while True:
                try:
                    if not os.path.exists(dir_atual.replace('P_','F_')):
                        os.replace(dir_atual,dir_atual.replace('P_','F_'))
                        break
                except:
                    print("\n\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n\n\nERRO NA HORA DE SALVAR O EXCEL. FECHE A PREVISÃO OU FATURA DESTE STORAGE\n\n\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n\n")
                    input('FECHE O EXCEL E APERTE ENTER')
        
        print('\n\n')


def menu():
    while True:
        try:
            mes = int(input('\n\nDIGITE O MÊS QUE DESEJA REALIZAR O FECHAMENTO (EX: 01): '))
            if mes not in range(1,13):
                print('         Digite um número de 1 a 12 para representar o mês!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')

    while True:
        try:
            ano_atual = date.today().year
            ano = int(input('\n\nDIGITE O ANO QUE DESEJA REALIZAR O FECHAMENTO (EX: 2021): '))
            if ano not in range(ano_atual-1,ano_atual+2):
                print(f'         Digite um ano entre {str(ano_atual-1)} e {str(ano_atual+1)} de 1 a 12 para representar o mês!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')
    
    return ano,mes

def menu_2():
    while True:
        try:
            tipo_virada = int(input(
            "\n\n\n*****************************************************************\n\n\
        1 - VIRAR STORAGES BLACKBIRD\n\n\
        2 - VIRAR STORAGES PLANILHA\n\n\
        3 - VIRAR STORAGES PRISMA\n\n\
        4 - VIRAR STORAGES LIBERTY\n\n\
        99 - VIRAR TODOS STORAGES PORTO (BLACKBIRD, PLANILHA E PRISMA)\n\n******************************************************************\nDigite a opção desejada: "))
            if tipo_virada not in [1,2,3,4,99]:
                print('         Selecione uma das opções oferecidas!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')

    return tipo_virada

def try_to_read_a_path(path_atual):
    try:
        df = pd.read_excel(path_atual,skiprows=3)
    except FileNotFoundError:
        new_path = path_atual.replace('P_','F_')
        df = pd.read_excel(new_path,skiprows=3)
    
    return df


if __name__ == '__main__':
    main()