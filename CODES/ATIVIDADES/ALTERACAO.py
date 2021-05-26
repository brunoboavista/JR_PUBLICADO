import pandas as pd
import datetime as dt
from dateutil import parser
from dateutil.relativedelta import relativedelta

def alteracao(df_atual,df_tabela,coberturas_possiveis):
    
    df_atual['BOX'] = df_atual['BOX'].astype(str)
    
    df_temporario = df_atual.loc[df_atual['STATUS']=='ATIVO']
    box_possiveis_de_alteracao = list(df_temporario['BOX'].astype(str))

    df_atual,continuar = processando_alteracao(df_atual,df_tabela,box_possiveis_de_alteracao,coberturas_possiveis)

    return df_atual,continuar


def processando_alteracao(df_atual,df_tabela,box_possiveis_de_alteracao,coberturas_possiveis):
    
    print('\n\n**************************************************************')
    pode_sair = False
    continuar = 'N'
    
    while True:
        box_escolhido = (input('Digite o box que deseja realizar a alteração: '))

        if box_escolhido not in box_possiveis_de_alteracao:
            print('         Informe um box válido para realizar a alteração\n')
        
        else:
            print('\n',df_atual.loc[df_atual['BOX']==box_escolhido],'\n')
            while True:
                confirma = input('Confirma este box (S ou N)? ').upper()
                if confirma not in ['S','N']:
                    print("   Digite apenas 'S' ou 'N'")
                elif confirma == 'N':
                    print('**************************************************************')
                    break
                else:
                    pode_sair = True
                    break
        if pode_sair == True:
            break
    print('\n***************************************************************\n')
    print('\nDIGITE OS DADOS A SEREM ALTERADOS, PARA NÃO ALTERAR DEIXE EM BRANCO\n')

    # BOX
    while True:
        box = input('   INFORME O BOX SEGURADO: ')
        pode_alterar = True
        if box == '':
            break
        elif box not in box_possiveis_de_alteracao:
            box_possiveis_de_alteracao.append(box)
            box_possiveis_de_alteracao.remove(box_escolhido)
            df_atual = arrumando_df(df_atual,df_tabela,'BOX',box_escolhido,box)
            break
        else: 
            print('     Este Box já está ocupado, favor cancelar o box ativo antes de prosseguir.')
            pode_alterar = False
            break
    
    if pode_alterar:
        # COBERTURA
        while True:
            try:
                incendio = input('   DIGITE A COBERTURA DE INCÊNDIO CONTRATADA: ')
                if incendio == '':
                    break
                elif int(incendio) in coberturas_possiveis:
                    df_atual = arrumando_df(df_atual,df_tabela,'INCÊNDIO',box_escolhido,int(incendio))
                    break
                else: 
                    print('     Digite uma cobertura válida para esta tabela!!!')
            except ValueError:
                    print('     Caso tenha digitado uma cobertura, use apenas números!!!')
        
        # CPF/CNPJ
        while True:
            cpf_cnpj = input('   DIGITE O CPF/CNPJ DO CLIENTE EM CASO DE ALTERAÇÃO: ')
            if cpf_cnpj == '':
                break
            else:
                confirma = input(f'     Confirma a alteração do CPF/CNPJ do segurado (S ou N)? ').upper()
                if confirma not in ['S','N']:
                    print("   Digite apenas 'S' ou 'N'")
                elif confirma == 'S':
                    df_atual = arrumando_df(df_atual,df_tabela,'CPF/CNPJ',box_escolhido,cpf_cnpj)
                    break
        
        # NOME
        while True:
            nome = input('   INFORME O NOME DO SEGURADO:').upper()
            if nome == '':
                break
            else:
                confirma = input(f'     Confirma a alteração do nome do segurado para {nome} (S ou N)? ').upper()
                if confirma not in ['S','N']:
                    print("   Digite apenas 'S' ou 'N'")
                elif confirma == 'S':
                    df_atual = arrumando_df(df_atual,df_tabela,'NOME',box_escolhido,nome)
                    break

        # PERGUNTANDO SE DESEJA CONTINUAR      
        while True:
            continuar = input('\n\nDeseja alterar mais clientes neste storage (S ou N): ').upper()
            if continuar not in ['S','N']:
                print("   Digite apenas 'S' ou 'N'")
            else:
                break

    return df_atual,continuar
            

def arrumando_df(df_atual,df_tabela,coluna,box_escolhido,new_value):
    
    df_atual.reset_index(inplace=True)
    df_atual.loc[df_atual['BOX']==box_escolhido, coluna] = new_value

    if coluna == 'INCÊNDIO':
        df_atual = df_atual.merge(df_tabela,on='INCÊNDIO')
        df_atual.columns = df_atual.columns.str.replace('_y','') 
        new_columns = [c for c in df_atual.columns if '_x' not in c]
        new_columns.remove('RECEBIMENTO JR')
        new_columns.append('RECEBIMENTO JR')
    else:
        new_columns = [c for c in df_atual.columns if '_x' not in c]
    
    new_columns.remove('index')
    df_atual.sort_values(by='index',inplace=True)
    df_atual = df_atual[new_columns]
    
    return df_atual
