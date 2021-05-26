import pandas as pd
import datetime as dt
from dateutil import parser
from datetime import datetime

def cancelamento(df_atual,sistema):

    # TRANSFORMA CAMPO DE BOX EM STRING
    df_atual['BOX'] = df_atual['BOX'].astype(str)

    # CRIA LISTA DE CANCELADOS E POSSÍVEIS DE SER CANCELADOS
    df_temporario = df_atual.loc[df_atual['STATUS']=='ATIVO']
    lista_to_cancel = list(df_temporario['BOX'])
    df_temporario = df_atual.loc[df_atual['STATUS']!='ATIVO']
    lista_canceled = list(df_temporario['BOX'])

    # CHAMANDO CANCELAMENTO
    df_atual,continuar = processando_cancelamento(df_atual,lista_to_cancel,lista_canceled,sistema)

    return df_atual,continuar


def processando_cancelamento(df_atual,lista_to_cancel,lista_canceled,sistema):
    
    print('\n\n**************************************************************')
    pode_sair = False
   
    # PERGUNTA O BOX E VALIDA SE PODE CONTINUAR PROCESSO
    while True:
        opcao_escolhida = (input('Digite o box que deseja realizar o cancelamento: '))
        
        if opcao_escolhida.upper() == 'SAIR':
            break

        elif opcao_escolhida not in lista_to_cancel:
            print('         Informe um box válido para cancelar\n')
        
        elif opcao_escolhida in lista_canceled:
            print('         O box informado já está cancelado\n')
        
        else:
            print('\n',df_atual.loc[df_atual['BOX']==opcao_escolhida],'\n')
            while True:
                confirma = input('Confirma a exclusão deste box (S ou N)? ').upper()
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
    print('\n***************************************************************\n\n')

    # FINALIZA PROCESSO, CASO TENHA SIDO APERTADO UM CANCELAMENTO A MAIS
    if opcao_escolhida.upper() != 'SAIR':
        
        # VALIDANDO DATA DE SAÍDA
        date = valida_data()
        df_atual.loc[(df_atual['BOX'] == opcao_escolhida),'RECEBIMENTO JR'] = date
        
        # MUDA O STATUS DEPENDENDO DA SEGURADORA E DIAS ATIVOS
        if sistema == 'LIBERTY':
            df_atual.loc[(df_atual['BOX'] == opcao_escolhida),'STATUS'] = 'CANCELADO'
        else:
            SAIDA_BOX_CANCELADO = date

            df_atual.loc[(df_atual['BOX'] == opcao_escolhida),'ENTRADA_CONSIDERAR'] = df_atual['ENTRADA'].dt.date
            df_atual.loc[(df_atual['BOX'] == opcao_escolhida),'DIF'] = (SAIDA_BOX_CANCELADO - df_atual['ENTRADA_CONSIDERAR']).dt.days
            
            # CRIANDO CAMPO DE DIFRENÇA DE DIAS ATIVOS
            df_atual.loc[(df_atual['BOX'] == opcao_escolhida) & (df_atual['DIF'] > 7),'STATUS'] = 'ÚLTIMA COBRANÇA'
            df_atual.loc[(df_atual['BOX'] == opcao_escolhida) & (df_atual['DIF'] <= 7),'STATUS'] = 'CANCELADO'
            df_atual.drop(['ENTRADA_CONSIDERAR','DIF'],axis=1,inplace=True)
            
        # VALIDA SE AINDA HÁ BOXES PARA SEREM CANCELADOS
        lista_to_cancel.remove(opcao_escolhida)
        if len(lista_to_cancel) == 0:
            print('\n\nSEM MAIS CLIENTES PARA CANCELAR')
            continuar = 'N'
        
        # PERGUNTA SE QUER CONTINUAR
        else:
            while True:
                continuar = input('\n\nDeseja cancelar mais clientes neste storage (S ou N): ').upper()
                if continuar not in ['S','N']:
                    print("   Digite apenas 'S' ou 'N'")
                else:
                    break
    
    # CASO NÃO DESEJE CANCELAR MAIS UM BOX
    else:
        continuar = 'N'

    
    return df_atual,continuar


def valida_data():

    while True:
        try:
            date = parser.parse(input("INSIRA A DATA DE RECEBIMENTO DA EXCLUSÃO (Ex: 01/01/2021): "),dayfirst=True).date()
        
            while True:
                confirma_data = input(f'Confirma a data {date.strftime("%d/%m/%Y")} (S ou N)? ').upper()
                
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
