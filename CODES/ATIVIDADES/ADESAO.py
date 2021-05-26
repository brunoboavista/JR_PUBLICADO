import pandas as pd
import datetime as dt
from dateutil import parser
from dateutil.relativedelta import relativedelta


def adesao(df_atual,df_tabela,coberturas_possiveis,sistema,seguradora,fechamento_previsao):
    
    # EXTRAI NOME DAS COLUNAS PARA USAR NO DF TEMPORARIO
    col_names = df_atual.columns
    
    while True:
        # INICIA VARIÁVEIS DO PROCESSO
        box_ocupados = list(df_atual['BOX'].astype(str))
        df_temporario = pd.DataFrame(columns=col_names)
        
        # INICIA CADASTRO
        print('\n\n*********************************************************')
        df_temporario = cadastro(df_temporario,box_ocupados,coberturas_possiveis,sistema,seguradora,fechamento_previsao)
        print('*********************************************************\n\n')
        
        # ARRUMA MERGE DOS PRÊMIOS
        df_temporario.dropna(axis=1,inplace=True)
        df_temporario = df_temporario.merge(df_tabela,on='INCÊNDIO')
        
        # DROP NAS COLUNAS DE DATA, CASO LIBERTY
        if sistema == 'LIBERTY':
            df_temporario.drop(columns=['ENTRADA','FIM DA VIG'],inplace=True)
        
        # CONFIRMAÇÃO DO BOX
        print(df_temporario)
        while True:
            confirma = input('\n     Confirma a inclusão deste box (S ou N)? ').upper()
            if confirma not in ['S','N']:
                print("   Digite apenas 'S' ou 'N'")
            else:
                break
        
        # CASO CONFIRMADA A INCLUSÃO
        if confirma == 'S':
            df_atual = df_atual.append(df_temporario,ignore_index=True)
            while True:
                continuar = input('\n\nDeseja inserir mais clientes neste storage (S ou N): ').upper()
                if continuar not in ['S','N']:
                    print("   Digite apenas 'S' ou 'N'")
                else:
                    break
            if continuar == 'N':
                break

        
        # CASO NÃO CONFIRMADA A INCLUSÃO
        else:
            salvar_adesoes_ja_feitas = input('\n     Deseja sair do menu e salvar as últimas adesões confirmadas (S ou N)? ').upper()
            if salvar_adesoes_ja_feitas not in ['S','N']:
                print("   Digite apenas 'S' ou 'N'")
            elif salvar_adesoes_ja_feitas == 'S':
                continuar = 'N'
                break
    
    return df_atual,continuar
    
def cadastro(df_temporario,box_ocupados,coberturas_possiveis,sistema,seguradora,fechamento_previsao):
    
    # BOX
    while True:
        box = input('   INFORME O BOX SEGURADO: ')
        if box == '':
            print('     DIGITE UM BOX!!!')
        elif box not in box_ocupados:
            break
        else:
            confirma = input('     Este Box já está ocupado no Storage escolhido, deseja incluir mesmo assim? (S ou N): ').upper()
            if confirma not in ['S','N']:
                print("   Digite apenas 'S' ou 'N'")
            elif confirma == 'S':
                break
    
    # COBERTURA
    while True:
        try:
            incendio = int(input('   DIGITE A COBERTURA DE INCÊNDIO CONTRATADA: '))
            if incendio in coberturas_possiveis:
                box_ocupados.append(box)
                break
            else: 
                print('     Digite uma cobertura válida para esta tabela!!!')
        except ValueError:
                print('     Digite uma cobertura!!!')

    # NOME
    while True:
        nome = input('   INFORME O NOME DO SEGURADO: ').upper()
        if nome == '':
            print('     Digite um nome válido!!!')
        else:
            box_ocupados.append(box)
            break

    # CPF/CNPJ
    while True:
        cpf_cnpj = input('   INFORME O CPF/CNPJ DO SEGURADO: ')
        if cpf_cnpj == '':
            print('     Digite um CPF/CNPJ válido!!!')
        else:
            break
    
    # DATA DE INICIO
    if sistema in ['BLACKBIRD','LIBERTY']:
        dt_ini = dt.date.today()
    else:
        dt_ini = valida_data()
    
    # DATA FIM DA VIGÊNCIA
    dt_fim = dt_ini + relativedelta(dt_ini,days=30)

    # STATUS
    status = 'ATIVO'

    # INSERINDO NOVO CLIENTE
    new_row = pd.Series(data={'NOME':nome,'INCÊNDIO':incendio,'BOX':box,'ENTRADA':dt_ini,'FIM DA VIG':dt_fim,'CPF/CNPJ':cpf_cnpj,'STATUS':status})
    df_temporario = df_temporario.append(new_row,ignore_index=True)

    # VALIDA SE CLIENTE TEM O FIM DE VIGÊNCIA DENTRO DO PERÍODO DE FATURAMENTO
    if seguradora == 'PORTO' and dt_fim == fechamento_previsao:
        new_data_fim = dt_fim + relativedelta(dt_ini,days=30)
        row_duplica_adesao_25 = pd.Series(data={'NOME':nome,'INCÊNDIO':incendio,'BOX':box,'ENTRADA':dt_fim,'FIM DA VIG':new_data_fim,'CPF/CNPJ':cpf_cnpj,'STATUS':status})
        df_temporario = df_temporario.append(row_duplica_adesao_25,ignore_index=True)

    return df_temporario

def valida_data():

    while True:
        try:
            date = parser.parse(input("   INSIRA A DATA DE INÍCIO DO CLIENTE (Ex: 01/01/2021): "),dayfirst=True).date()
            
            while True:
                confirma_data = input(f'    Confirma a data {date.strftime("%d/%m/%Y")} (S ou N)? ').upper()
                
                if confirma_data not in ['S','N']:
                    print("   Digite apenas 'S' ou 'N'")
                
                elif confirma_data == 'S':
                    return date
                
                else:
                    break

        except KeyboardInterrupt:
            break
            
        except:
            print('     Verifique a data e tente novamente!!')
