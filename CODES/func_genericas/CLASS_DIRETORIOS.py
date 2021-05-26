from dateutil.relativedelta import relativedelta
from datetime import date
import glob
import os
import time

class DIRETORIOS(object):

    def __init__(self):

        # DEFINE DATA DE HOJE
        today = date.today().strftime('%Y_%m_%d')

        # DEFINE CAMINHOS PADRÕES
        self.root = '//SRVJRMARTO/Files/(1) SELF STORAGE 2020-21/'
        self.dir_bases = self.root + 'STORAGES/'
        self.dir_boletos = self.root + 'ALFA/BOLETOS/'
        self.lista_storages = self.dir_bases + '1.0 APOIO/LISTA DE STORAGES.xlsx'
        self.path_movimentacoes_do_dia = self.dir_bases + '1.0 MOVIMENTACOES/MOVIMENTACOES_' + today + '.xlsx'


    # DEFINE TUDO USADO NA MOVIMENTAÇÃO
    def definindo_path_atual(self,seguradora,razao_social,nome_fantasia,tabela,sistema):
        
        diretorio_tabelas = self.dir_bases + '1.0 TABELAS/' + seguradora + '/TABELA_'

        # DEFINE CAMINHO DA PASTA DO STORAGE
        main_root = self.dir_bases + seguradora + '/' + nome_fantasia + '/'

        # DEFINE O DIRETÓRIO DE PREVISÃO MAIS RECENTE
        list_previsoes = glob.glob(main_root + 'P_*')
        if len(list_previsoes)!=0:
            most_recent_dir = min(list_previsoes) + '/'

        else:
            # MARCA MES_ANO DO PRÓXIMO DIA 25, A PARTIR DE HOJE
            today = date.today()
            nex_day_25 = today + relativedelta(day=25)
            ano_mes = nex_day_25.strftime('P_%Y_%m')

            # CRIA DIRETÓRIO MAIS RECENTE, CASO NÃO HAJA
            most_recent_dir = main_root + ano_mes + '/'
        
        # DEFINE DIRETÓRIOS DO STORAGE
        self.path_atual = most_recent_dir + razao_social + '.xlsx'
        self.path_tabela = diretorio_tabelas + tabela + '.xlsx'
        self.path_protocolo = main_root + '1.0 PROTOCOLOS/'     


    # DEFINE RÓTULOS DO STORAGE
    def definindo_labels(self,razao_social,tabela,susep,estipulante,modalidade,endereco,seguradora):
        self.label_1 = razao_social
        if seguradora == 'LIBERTY':
            self.label_2 = f'CARTA OFERTA: {str(susep)}'
        else:
            self.label_2 = f'SUSEP: {str(susep)} - ESTIPULANTE: {str(estipulante)} - MODALIDADE: {str(modalidade)} - CÓDIGO DE OPERAÇÃO: {str(tabela)}'
        self.label_3 = endereco


    # DEFINE TUDO USADO NA HORA DE VIRAR
    def definindo_virada(self,seguradora,razao_social,nome_fantasia,ano_mes_fechamento,ano_mes_virada):
        main_root = self.dir_bases + seguradora + '/' + nome_fantasia + '/'
        
        most_recent_dir = main_root + ano_mes_fechamento + '/'
        self.path_atual = most_recent_dir + razao_social + '.xlsx'
        self.dir_fechamento = most_recent_dir
        
        dir_virada = main_root + ano_mes_virada + '/'
        self.path_virada = dir_virada + razao_social + '.xlsx'
        self.dir_virada = dir_virada


    # DEFINE TUDO USADO NOS ENVIOS
    def definindo_envios(self,seguradora,razao_social,nome_fantasia,ano_mes_fechamento,vcto_boleto,susep,tipo_envio):
        
        # DEFINE MAIN ROOT
        main_root = self.dir_bases + seguradora + '/' + nome_fantasia + '/'
        
        # DEFINE ANO E MES, COMO IRÁ NO ASSUNTO DO EMAIL
        ano_mes_subjetc = ano_mes_fechamento[-2:] + '/' + ano_mes_fechamento[2:6]

        # DEFINE CAMINHO DO PROTOCOLO
        list_protocolos = glob.glob(main_root + '1.0 PROTOCOLOS/*.xlsx')
        if len(list_protocolos) > 0 :
            self.path_protocolo = max(list_protocolos)
        else: 
            self.path_protocolo = None

        # CRIA RACIONAIS DEPENDENDO DO TIPO DE ENVIO
        if tipo_envio == 1:
            ano_mes_fechamento = ano_mes_fechamento.replace('P_','F_')
            dir_to_send = main_root + ano_mes_fechamento + '/'
            self.path_to_send = dir_to_send + razao_social + '.xlsx'
            self.subject = f'FECHAMENTO DE FATURA {seguradora} | {razao_social} | {ano_mes_subjetc}'

        elif tipo_envio in [2,99]:
            today = date.today().strftime("%d/%m/%Y")
            dir_to_send = main_root + ano_mes_fechamento + '/'
            self.path_to_send = dir_to_send + razao_social + '.xlsx'
            self.subject = f'PREVISÃO DA FATURA {seguradora} | {razao_social} | {today}'

        elif tipo_envio in [3,4]:
            ano_mes_fechamento = ano_mes_fechamento.replace('P_','F_')
            ano = vcto_boleto[2:6]
            dir_to_send = main_root + ano_mes_fechamento + '/'
            self.path_to_send = dir_to_send + razao_social + '.xlsx'
            self.subject = f'FATURA FECHADA {seguradora} | {razao_social} | {ano_mes_subjetc}'
            self.path_boleto = self.dir_boletos + 'BOLETOS ' + ano+ '/' + vcto_boleto + '/' + seguradora + '/' + 'CONFERIDOS/QUIVER/'  + nome_fantasia + '.pdf'

        elif tipo_envio == 5:
            dir_to_send = main_root + ano_mes_fechamento + '/'
            self.path_to_send = dir_to_send + razao_social + '.xlsx'
            self.subject = f'PREVISÃO DE FATURA {seguradora} | {razao_social} | {ano_mes_subjetc}'

if __name__ == '__main__':
    return_class = DIRETORIOS()