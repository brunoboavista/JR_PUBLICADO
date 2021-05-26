import pandas as pd
from dateutil.relativedelta import relativedelta

def main(path_atual,path_virada,dir_virada,fechamento_virada):

    # LENDO EXCEL A SER FECHADO
    try:
        df_atual = pd.read_excel(path_atual,skiprows=3)
    except FileNotFoundError:
        new_path = path_atual.replace('P_','F_')
        df_atual = pd.read_excel(new_path,skiprows=3)
    
    # REMOVENDO LINHAS EM BRANCO E LINHA DE TOTAIS
    df_atual = df_atual.loc[df_atual['BOX'] != 'TOTAIS'].copy()
    df_atual.dropna(subset=['BOX'],axis=0,inplace=True)

    if not df_atual.empty:
        # ALTERANDO STATUS DAS ÚLTIMAS COBRANÇAS
        df_atual.loc[df_atual['STATUS']=='ÚLTIMA COBRANÇA','STATUS'] = 'CANCELADO'

        # REMOVENDO VIGÊNCIAS DUPLICADAS DEVIDO AO DIA 25 DO MÊS ANTERIOR
        df_atual.sort_values(by='ENTRADA',inplace=True)
        df_atual.drop_duplicates(subset=['BOX','STATUS','NOME','CPF/CNPJ'],keep='last',inplace=True,ignore_index=True)

        # VIRANDO CLIENTES COM SEGUNDA VIGÊNCIA DIA 25, CASO HAJA
        df_clientes_25 = df_atual.loc[(df_atual['FIM DA VIG'].dt.day == 25) & (df_atual['ENTRADA'].dt.day == 26) & ~(df_atual['ENTRADA'].dt.date > fechamento_virada)].copy()
        df_clientes_25 = virando_datas(df_clientes_25,fechamento_virada)
        df_clientes_25.loc[df_clientes_25['STATUS']=='ATIVO','ENTRADA'] = df_clientes_25['ENTRADA'] - pd.DateOffset(days=30)
        df_clientes_25.loc[df_clientes_25['STATUS']=='ATIVO','FIM DA VIG'] = df_clientes_25['FIM DA VIG'] - pd.DateOffset(days=30)

        # VIRANDO DATAS DOS CLIENTS ATUAIS
        df_atual = virando_datas(df_atual,fechamento_virada)
        
        # CHECANDO CLIENTES COM VIGÊNCIA INICIAL MAIOR QUE DIA 25 DO MÊS DE FECHAMENTO DA VIRADA
        df_clientes_mais_25 = df_atual.loc[df_atual['ENTRADA'].dt.date > fechamento_virada].copy()
        if not df_clientes_mais_25.empty:
            print('A nova previsão possui clientes fora do período de faturamento da seguradora e deve ser revisada!!!')

        # JUNTANDO DATAFRAME PADRÃO E DF CLIENTES DO DIA 25
        df_virada = df_atual.append(df_clientes_25,ignore_index=True)
    else:
        df_virada = df_atual.copy()
    
    df_virada.sort_values(by='ENTRADA',inplace=True)
    return df_virada


def virando_datas(df,fechamento_virada):
    inicio_vigencia = fechamento_virada - relativedelta(months=1) + relativedelta(days=1)
    df.loc[(df['STATUS']=='ATIVO') & ~(df['ENTRADA'].dt.date > inicio_vigencia),'ENTRADA'] = df['FIM DA VIG']
    df.loc[df['STATUS']=='ATIVO','FIM DA VIG'] = df['ENTRADA'] + pd.DateOffset(days=30)

    return df