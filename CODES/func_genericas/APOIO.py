import pandas as pd
import FORMATANDO_EXCEL_CONVERTER
import FORMATANDO_EXCEL
import glob
from unidecode import unidecode
import os
import shutil
import numpy as np
import time

def convert_to_new():
    merdinha = 'n'

    my_paths = glob.glob('//SRVJRMARTO/Files/(1) SELF STORAGE 2020-21/STORAGES/LIBERTY/*/P_2021_04/*.xlsx')
    # my_paths = glob.glob('C:/Users/Bruno/Desktop/JR/CONVERSAO/1. LIBERTY/*/*2021_04/*.xlsx')

    print()
    for path in my_paths:
        if '~$' not in path and '(imobiliário)':
            
            print('\n\n\n\n',path,'\n')
            rodar = input('Processo Pausado para conferência, aperte enter para continuar ').upper()
            merdinha = 'n'
            if rodar == 'S':
            # if True:
                # DEFINE DIRETÓRIO DE SAÍDA
                path_quebrado = path.replace('\\','/').split('/')
                name = unidecode(path_quebrado[-1].upper()).replace('XLSX','xlsx')
                dir_exit = '/'.join(path_quebrado[:-1]) + '/'
                destino = dir_exit + '1.0 ' + name

                # time.sleep(60)
                # COPIA OS LABELS DO ARQUIVO JÁ EXISTENTE
                df_labels = pd.read_excel(path,skiprows=34)
                labels = list(df_labels['Unnamed: 0'])[:4]
                labels = [str(label).replace(str(np.nan),'') for label in labels]

                # AJUSTE PARA CONVERTER LIBERTY
                if 'BOX' in labels[-1] or 'CARTA-OFERTA' in labels[-1]:
                    labels = labels[:-1]
                else:
                    labels = [labels[0], 'CARTA OFERTA: ATIVA', labels[3]]
                # print('\n\n')
                # print(labels[0])
                # print(labels[1])
                # print(labels[2])


                # LÊ O ARQUIVO PARA TRATÁ-LO
                df_atual = pd.read_excel(path,skiprows=38)

                if 'BOX' not in list(df_atual.columns):
                    df_atual = pd.read_excel(path,skiprows=39)

                # VERIFICANDO AS COLUNAS UTILIZÁVEIS
                try:
                    columns_to_consider = [c for c in list(df_atual.columns) if 'Unnamed' not in c and 'CARTA-OFERTA' not in c]

                except:
                    print('\n\n',path,'\n\n')

                # LIMPANDO DATA FRAME
                df_atual = df_atual[columns_to_consider]
                df_atual = df_atual.loc[(df_atual['BOX'] != 'Total')]
                df_atual = df_atual.loc[~(df_atual['BOX'].isna())]
                
                # COLETANDO DADOS ANTIGOS
                cols_antigas = list(df_atual.columns)
                shape_antigo = df_atual.shape
                # df_atual['ENTRADA'] = pd.to_datetime(df_atual['ENTRADA'],dayfirst=True)
                
                # # RENOMEIA AS COLUNAS
                new_cols = [c.replace('  ',' ').replace(' PORTO SEGURO','').replace('INCENDIO','INCÊNDIO').replace('RECEBIM. JR.','RECEBIMENTO JR').replace('HIDRÁULICO','HIDRÁULICOS').replace('HIDRÁULICOS','ÁGUA') for c in df_atual.columns]
                df_atual.columns = new_cols
                
                # MARCA COMO TEMPO E TIRA CARACTERES DOS NOMES
                df_atual['RECEBIMENTO JR'] = pd.to_datetime(df_atual['RECEBIMENTO JR'])
                df_atual['INCÊNDIO'] = df_atual['INCÊNDIO'].astype(int)
                try:
                    df_atual['NOME'] = df_atual['NOME'].str.upper()
                    df_atual['NOME'] = df_atual['NOME'].apply(unidecode)
                except AttributeError:
                    print('DEU MERDA NA TABELA')
                    merdinha = 's'

                if not merdinha == 's':
                    # CONFERE SHAPE E COLUNAS
                    print(cols_antigas)
                    # print(list(df_atual.columns))
                    print(shape_antigo)
                    # print(df_atual.shape)
                        

                    # # CRIANDO CAMPO DE DIFRENÇA DE DIAS ATIVOS
                    # if 'ENTRADA' in df_atual.columns:
                    #     df_atual['DIAS ATIVOS'] = (df_atual['RECEBIMENTO JR'] - df_atual['ENTRADA']).dt.days
                    #     # MARCANDO SE É CANCELAMENTO OU ÚLTIMA COBRANÇA
                    #     df_atual.loc[(~df_atual['RECEBIMENTO JR'].isna()) & (df_atual['STATUS'] == 'ATIVO'),'RECEBIMENTO JR'] = pd.NaT
                    #     df_atual.loc[(~df_atual['RECEBIMENTO JR'].isna()) & (df_atual['DIAS ATIVOS'] > 7),'STATUS'] = 'ÚLTIMA COBRANÇA'
                    #     df_atual.loc[(~df_atual['RECEBIMENTO JR'].isna()) & (df_atual['DIAS ATIVOS'] <= 7),'STATUS'] = 'CANCELADO'
                    #     df_atual.drop(['DIAS ATIVOS'],axis=1,inplace=True)
                    
                    # SE NÃO QUISER CONSIDERAR OS CANCELADOS ANTIGOS
                    df_atual.loc[(~df_atual['RECEBIMENTO JR'].isna()) & (df_atual['STATUS'] == 'ATIVO'),'RECEBIMENTO JR'] = pd.NaT
                    FORMATANDO_EXCEL_CONVERTER.main(destino,path,df_atual,labels[0],labels[1],labels[2])

                    # CRIANDO DIRETÓRIO DO SISTEMA MANUAL
                    dir_antigo = dir_exit + 'ANTIGO/'
                    if not os.path.exists(dir_antigo):
                        os.makedirs(dir_antigo)
                    # shutil.copy(path,dir_antigo + name)
                    shutil.move(path,dir_antigo + name)
                    os.rename(destino,destino.replace('1.0 ',''))
        

def pastas():
    my_paths = glob.glob('//SRVJRMARTO/Files/(1) SELF STORAGE 2020-21/STORAGES/PORTO/*/F*/')
    for pastas in my_paths:
        my_dir = pastas.replace('\\','/')
        new_name = pastas.replace('\\','/').replace('F 2021_02','F_2021_02')
        print(my_dir)
        print(new_name)
        # os.rename(pastas,new_name)
        # time.sleep(60)
        # print(pastas.replace('\\','/').split('/')[-1])

def arquivos():
    my_paths = glob.glob('//SRVJRMARTO/Files/(1) SELF STORAGE 2020-21/STORAGES/LIBERTY/*/*2021_04/*.xlsx')
    for pastas in my_paths:
        if  '~$' not in pastas:
            print(pastas.replace('\\','/').replace('.xlsx','').split('/')[-3])


import xlwings as xw
def open_():
        my_paths = glob.glob('//SRVJRMARTO/Files/(1) SELF STORAGE 2020-21/STORAGES/LIBERTY/*/P_2021_05/*.xlsx')
        for pastas in my_paths:
            # print(pastas)
            if  '~$' not in pastas:
                my_dir = pastas.replace('\\','/')
                FANTASIS = my_dir.split('/')[-3]
                # print('\n\n\n\n',pastas,'\n')
                df_atual = pd.read_excel(pastas,skiprows=3,sheet_name=1)
                # if 'DANOS HIDRÁULICO' in list(df_atual.columns):
                #     xw.Book(pastas)
                #     print(FANTASIS,'&',list(df_atual.columns))
                print(FANTASIS)
                print(list(df_atual.columns))



if __name__ == "__main__":
    # pastas()
    # arquivos()
    # convert_to_new()
    # print('descomente o código')
    open_()