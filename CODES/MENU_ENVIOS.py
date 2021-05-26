from datetime import date
import os
import pandas as pd
import MAIL.ENVIA_EMAIL as ENVIA_EMAIL
import func_genericas.CLASS_DIRETORIOS as CLASS_DIRETORIOS
from dateutil.relativedelta import relativedelta
import shutil

def main():

    # CHAMANDO MENUS
    tipo_envio = menu()
    ano,mes = menu_2()

    # DEFININDO VARIAVEIS DE DATA
    ano_mes_fechamento = date(ano,mes,1).strftime('P_%Y_%m')
    vcto_boleto = (date(ano,mes,1) + relativedelta(months=1)).strftime('V_%Y_%m')
    data_corte = date(
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

    # LENDO EMAILS PORTO
    df_emails_porto = pd.read_excel(path_todos_storages,sheet_name='CC')
    mail_porto_to = list(df_emails_porto['EMAIL TO PORTO'])[0]
    mail_porto_cc = list(df_emails_porto['EMAIL CC PORTO'])[0]

    # DAR UM LOC NOS CRITÉRIOS DE SEGURADORA, SISTEMA
    if tipo_envio in [1,5]:
        storages = storages.loc[storages['SEGURADORA']=='PORTO']
        storages = storages.loc[storages['SISTEMA']!='BLACKBIRD']
    
    elif tipo_envio == 2:
        list_of_today = list(pd.read_excel(return_class.path_movimentacoes_do_dia)['STORAGE'])
        # storages = storages.loc[storages['SEGURADORA']=='PORTO']
        storages['CHAVE'] = storages['SEGURADORA'] + '_' + storages['RAZÃO SOCIAL']
        storages = storages.loc[storages['CHAVE'].isin(list_of_today)]
    
    elif tipo_envio == 3:
        storages = storages.loc[storages['SEGURADORA']=='PORTO']
    
    elif tipo_envio == 4:
        storages = storages.loc[storages['SEGURADORA']=='LIBERTY']
    
    # CRIANDO LISTAS DE RAZÃO E FANTASIA  
    susep = list(storages['SUSEP/C.O'].astype(str))
    seguradora = list(storages['SEGURADORA'])
    razao_social = list(storages['RAZÃO SOCIAL'])
    nome_fantasia = list(storages['NOME FANTASIA'])
    email_to_protocolo = list(storages['PROTOCOLO E-MAILS TO'].fillna(''))
    email_cc_protocolo = list(storages['PROTOCOLO E-MAILS CC'].fillna(''))
    email_to_financeiro = list(storages['FINANCEIRO E-MAILS TO'].fillna(''))
    email_cc_financeiro = list(storages['FINANCEIRO E-MAILS CC'].fillna(''))

    print("\n\nINICIANDO ENVIOS...")
    print('***************************************************************************************************')
    
    # DEFINE VARIÁVEIS DOS ENVIOS, CONFORME TIPO E CLIENTE
    for i in range(1,len(nome_fantasia)+1):
        mail_to_protocolo = email_to_protocolo[i-1]
        mail_cc_protocolo = email_cc_protocolo[i-1]
        mail_to_financeiro = email_to_financeiro[i-1]
        mail_cc_financeiro = email_cc_financeiro[i-1]
        susep_selected = susep[i-1]
        seguradora_selected = seguradora[i-1]
        razao_social_selected = razao_social[i-1]
        nome_fantasia_selected = nome_fantasia[i-1]

        # CHAMANDO CLASSE PARTE 2
        return_class.definindo_envios(seguradora_selected, razao_social_selected, nome_fantasia_selected, ano_mes_fechamento, vcto_boleto, susep_selected, tipo_envio)
        subject = return_class.subject
        path_anexo = return_class.path_to_send
        path_protocolo = return_class.path_protocolo
        
        # ACRESCENTA BOLETO AO ENVIO
        if tipo_envio in [3,4]:
            path_boleto = return_class.path_boleto
            dir_enviados = '/'.join(path_boleto.split('/')[:-1]) + '/ENVIADOS/'
        
        # FECHAMENTO PARA PORTO
        if tipo_envio == 1:
            if os.path.exists(path_anexo):
                vazio = valida_se_vazio(path_anexo)
                if not vazio:
                    no_prazo = valida_se_no_prazo(path_anexo,data_corte)
                    if no_prazo:
                        print(f'ENVIANDO E-MAIL PARA FECHAMENTO: {nome_fantasia_selected}\n')
                        ENVIA_EMAIL.send_email(path_anexo, mail_porto_to, mail_porto_cc, subject, tipo_envio, path_protocolo)
                    else:
                        print(f'ENVIO {nome_fantasia_selected} NÃO REALIZADO, POIS A BASE POSSUI CLIENTES COM INÍCIO DE VIGÊNCIA APÓS A DATA DE CORTE\n')
                else:
                    print(f'ENVIO {nome_fantasia_selected} NÃO REALIZADO, POIS A BASE ESTÁ VAZIA\n')
            else:
                print(f'O STORAGE {nome_fantasia_selected} NÃO POSSUI FATURA FECHADA PARA O PERÍODO SELECIONADO\n')
        
        # ENVIAR PREVISÕES PARA PORTO
        if tipo_envio == 5:
            if os.path.exists(path_anexo):
                vazio = valida_se_vazio(path_anexo)
                if not vazio:
                    print(f'ENVIANDO FPREVISÃO DO STORAGE {nome_fantasia_selected} PARA A PORTO\n')
                    ENVIA_EMAIL.send_email(path_anexo, mail_porto_to, mail_porto_cc, subject, tipo_envio, path_protocolo)
                else:
                    print(f'ENVIO {nome_fantasia_selected} NÃO REALIZADO, POIS A BASE ESTÁ VAZIA\n')
            else:
                print(f'O STORAGE {nome_fantasia_selected} NÃO POSSUI PREVISÃO PARA O PERÍODO SELECIONADO\n')

        # ENVIAR PREVISÕES PARA OS CLIENTES DO DIA/TODOS OS CLIENTES
        if tipo_envio in [2,99]:
            if os.path.exists(path_anexo):
                if path_protocolo == None:
                    print(f'REALIZANDO ENVIO DO STORAGE SEM PROTOCOLO: {nome_fantasia_selected}\n')
                else:
                    print(f'REALIZANDO ENVIO DO STORAGE + ÚLTIMO PROTOCOLO: {nome_fantasia_selected}\n')
                ENVIA_EMAIL.send_email(path_anexo, mail_to_protocolo, mail_cc_protocolo, subject, tipo_envio, path_protocolo)
            else:
                print(f'O STORAGE {nome_fantasia_selected} NÃO POSSUI PREVISÃO PARA O PERÍODO SELECIONADO\n')

        # ENVIAR FECHAMENTOS + BOLETOS AOS CLIENTES PORTO/LIBERTY
        if tipo_envio in [3,4]:
            if os.path.exists(path_anexo):
                vazio = valida_se_vazio(path_anexo)
                if not vazio:
                    if os.path.exists(path_boleto):
                        print(f'REALIZANDO ENVIO DA FATURA FECHADA: {nome_fantasia_selected}\n')
                        ENVIA_EMAIL.send_email(path_anexo, mail_to_financeiro, mail_cc_financeiro, subject, tipo_envio, path_boleto)
                        
                        # MOVE BOLETOS PARA PASTA DE ENVIADOS
                        if not os.path.exists(dir_enviados):
                            os.makedirs(dir_enviados)
                        shutil.move(path_boleto,dir_enviados)
                    
                    elif os.path.exists(dir_enviados + path_boleto.split('/')[-1]):
                        print(f'BOLETO DO STORAGE {nome_fantasia_selected} JÁ FOI ENVIADO\n')

                    else:
                        print(f'NÃO HÁ BOLETO DO STORAGE {nome_fantasia_selected} NA PASTA\n')

                else:
                    print(f'ENVIO {nome_fantasia_selected} NÃO REALIZADO, POIS A BASE ESTÁ VAZIA\n')
            else:
                print(f'O STORAGE {nome_fantasia_selected} NÃO POSSUI FATURA FECHADA PARA O PERÍODO SELECIONADO\n')
        
    print('\nFIM\n')


def menu_2():
    while True:
        try:
            mes = int(input('\n\nDIGITE O MÊS DO FECHAMENTO (EX: 01): '))
            if mes not in range(1,13):
                print('         Digite um número de 1 a 12 para representar o mês!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')

    while True:
        try:
            ano_atual = date.today().year
            ano = int(input('\n\nDIGITE O ANO DO FECHAMENTO (EX: 2021): '))
            if ano not in range(ano_atual-1,ano_atual+2):
                print(f'         Digite um ano entre {str(ano_atual-1)} e {str(ano_atual+1)} de 1 a 12 para representar o mês!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')

    while True:
        print(f'\n\nO PROCESSO INICIARÁ COM ENVIOS REFERENTE AO FECHAMENTO DE {str(mes).zfill(2)}/{ano}')
        last_chance = input('\n     PARA CONFIRMAR APERTE "ENTER", PARA CANCELAR APERTE CTRL+C')
        if last_chance == '':
            break
    
    return ano,mes

def menu():
    while True:
        try:
            tipo_envio = int(input(
            "\n*****************************************************************\n\n\
        1 - ENVIAR FECHAMENTO PARA PORTO\n\n\
        2 - ENVIAR PREVISÕES PARA CLIENTES DO DIA\n\n\
        3 - ENVIAR FECHAMENTOS + BOLETOS PARA OS CLIENTES PORTO\n\n\
        4 - ENVIAR FECHAMENTOS + BOLETOS PARA OS CLIENTES LIBERTY\n\n\
        5 - ENVIAR PREVISÕES PARA PORTO SEGURO\n\n\
        99 - ENVIAR PREVISÕES PARA TODOS OS CLIENTES\n\n******************************************************************\nDigite a opção desejada: "))
            if tipo_envio not in [1,2,3,4,5,99]:
                print('         Selecione uma das opções oferecidas!!')
            else:
                break
        except ValueError:
                print('\n\n         Digite um número!!')

    return tipo_envio

def valida_se_vazio(path_anexo):
    df = pd.read_excel(path_anexo,skiprows=3)
    df = df.loc[df['BOX'] !=  'TOTAIS']
    df.dropna(subset=['BOX'],axis=0,inplace=True)

    return df.empty

def valida_se_no_prazo(path_anexo,data_corte):
    df = pd.read_excel(path_anexo,skiprows=3)
    df = df.loc[df['BOX'] !=  'TOTAIS']
    df.dropna(subset=['BOX'],axis=0,inplace=True)
    df_clientes_mais_25 = df.loc[df['ENTRADA'].dt.date > data_corte].copy()

    return df_clientes_mais_25.empty



if __name__ == '__main__':
    main()





