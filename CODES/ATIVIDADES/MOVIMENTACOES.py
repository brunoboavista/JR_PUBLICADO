from ATIVIDADES.ADESAO import adesao
from ATIVIDADES.ALTERACAO import alteracao
from ATIVIDADES.CANCELAMENTO import cancelamento
import pandas as pd
import os
from datetime import date

def main(path_xlsx,path_tabela,fechamento_previsao,tipo_movimentacao,sistema,seguradora,*path_protocolo):

    # LENDO TABELA DE VALORES COBRADOS E CRIANDO COBERTURAS POSSÍVEIS
    df_tabela = pd.read_excel(path_tabela)
    coberturas_possiveis = list(df_tabela['INCÊNDIO'])
    
    if sistema == 'PRISMA':
        df_atual = it_is_prisma(path_protocolo[0],df_tabela,path_xlsx)
        if seguradora == 'LIBERTY':
            df_atual.drop(columns=['ENTRADA','FIM DA VIG'],inplace=True)

    else:
        # LENDO BASE DE CLIENTES ATUAIS
        columns_df = ['BOX','STATUS','NOME','CPF/CNPJ','ENTRADA','FIM DA VIG','INCÊNDIO'] + list(df_tabela.columns)[1:] + ['RECEBIMENTO JR']
        if not os.path.exists(path_xlsx):
            df_atual = pd.DataFrame(columns=columns_df)
        else:
            df_atual = pd.read_excel(path_xlsx,skiprows=3)
            df_atual = df_atual.loc[df_atual['BOX'] != 'TOTAIS']
            df_atual.dropna(subset=['BOX'], inplace=True)
            
            # FILTRANDO COLUNAS MESTRAS
            colunas_principais = [c for c in list(df_atual.columns) if c in ['BOX','STATUS','NOME','CPF/CNPJ','ENTRADA','FIM DA VIG','INCÊNDIO','RECEBIMENTO JR']]
            df_atual = df_atual[colunas_principais]

            # MERGE COM TABELA ATUAL
            df_atual = df_atual.merge(df_tabela,on='INCÊNDIO',how='left')

            # REORDENANDO COLUNAS
            if seguradora != 'LIBERTY':
                ordem_colunas = ['BOX','STATUS','NOME','CPF/CNPJ','ENTRADA','FIM DA VIG','INCÊNDIO'] + list(df_tabela.columns)[1:] + ['RECEBIMENTO JR']
            else:
                ordem_colunas = ['BOX','STATUS','NOME','CPF/CNPJ','INCÊNDIO'] + list(df_tabela.columns)[1:] + ['RECEBIMENTO JR']
            df_atual = df_atual[ordem_colunas]

        # CRIANDO ADESÃO
        if tipo_movimentacao == 1:
            while True:
                df_atual,continuar = adesao(df_atual,df_tabela,coberturas_possiveis,sistema,seguradora,fechamento_previsao)
                if continuar == 'N':
                    break
            print('\n\n********************************************\nTODAS AS INCLUSÕES FORAM REALIZADAS!!!')

        if df_atual.empty:
            print('\n\nNão é Possível realizar ALTERAÇÕES/CANCELAMENTO, pois a base está sem clientes!!!!')
        
        else:
            # CRIANDO ALTERAÇÃO
            if tipo_movimentacao == 2:
                while True:
                    df_atual,continuar = alteracao(df_atual,df_tabela,coberturas_possiveis)
                    if continuar == 'N':
                        break
                print('\n\n********************************************\nTODAS AS ALTERAÇÕES FORAM REALIZADAS!!!')

            # CRIANDO CANCELAMENTO
            if tipo_movimentacao == 3:
                while True:
                    df_atual,continuar = cancelamento(df_atual,sistema)
                    if continuar == 'N':
                        break
                print('\n\n********************************************\nTODAS AS EXCLUSÕES FORAM REALIZADAS FORAM REALIZADAS!!!')

    
    # AJUSTA COLUNAS DE DATAS CASO PORTO OU LIBERTY
    if seguradora != 'LIBERTY':
        
        # df_atual['RECEBIMENTO JR'] = df_atual['RECEBIMENTO JR'].fillna(pd.NaT)
        df_atual['ENTRADA'] = pd.to_datetime(df_atual['ENTRADA'],errors='coerce')
        df_atual['FIM DA VIG'] = pd.to_datetime(df_atual['FIM DA VIG'],errors='coerce')
        df_atual['RECEBIMENTO JR'] = pd.to_datetime(df_atual['RECEBIMENTO JR'],errors='coerce')

    return df_atual


def it_is_prisma(path_protocolo,df_tabela,path_xlsx):

    # LENDO PROTOCOLO
    df_protocolo = pd.read_excel(path_protocolo,skiprows=12)
    print('\n\nLENDO BASE DE CLIENTES PRISMA')
    
    # RENOMEANDO COLUNAS PROTOCOLO
    df_protocolo.rename(inplace=True,
            columns={
                'CUSTO TABELA':'PRÊMIO',
                'NOME DO CLIENTE':'NOME',
                'INÍCIO':'ENTRADA',
                'CPF':'CPF/CNPJ',
                'NÚMERO DO BOX':'BOX',
                'FIM':'FIM DA VIG',
                'TIPO':'STATUS'
            })
    
    # DEFINE DATA CORTE
    previsao = path_xlsx.replace('\\','/').split('/')[-2]
    ano_mes = previsao.split('_')
    data_corte = date(year=int(ano_mes[1]),month=int(ano_mes[2]),day=25)

    # ARRUMANDO CAMPOS DE DATAS
    df_protocolo['ENTRADA'] = pd.to_datetime(df_protocolo['ENTRADA'],dayfirst=True)
    df_protocolo['FIM DA VIG'] = pd.to_datetime(df_protocolo['FIM DA VIG'],dayfirst=True)
    df_protocolo.loc[df_protocolo['ENTRADA'].dt.date > data_corte,'ENTRADA'] = data_corte
    
    # MERGE BASES PROTOCOLO
    df_protocolo = df_protocolo.merge(df_tabela,on='PRÊMIO',how='outer')

    # AJUSTANDO COLUNAS
    df_protocolo['RECEBIMENTO JR'] = ''
    df_protocolo.drop('ENDEREÇO',inplace=True,axis=1)
    df_protocolo.dropna(subset=['BOX'],inplace=True)
    df_protocolo.loc[df_protocolo['STATUS']!='C','STATUS'] = 'ATIVO'
    df_protocolo.loc[df_protocolo['STATUS']=='C','STATUS'] = 'CANCELADO'

    # ARRUMANDO ORDEM DAS COLUNAS
    columns_df = ['BOX','STATUS','NOME','CPF/CNPJ','ENTRADA','FIM DA VIG','INCÊNDIO'] + list(df_tabela.columns)[1:] + ['RECEBIMENTO JR']
    df_protocolo = df_protocolo[columns_df]
    df_protocolo.sort_values(by='ENTRADA',inplace=True)

    print('\n\nBASE ATUALIZADA COM SUCESSO!!!\n\n')

    return df_protocolo
