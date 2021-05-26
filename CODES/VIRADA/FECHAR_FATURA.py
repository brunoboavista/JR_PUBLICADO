import pandas as pd

def main(path_atual,fechamento_virada):

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
        # CHECANDO CLIENTES COM VIGÊNCIA INICIAL MAIOR QUE DIA 25 DO MÊS DE FECHAMENTO DA VIRADA
        df_atual.loc[df_atual['ENTRADA'].dt.date > fechamento_virada,'STATUS'] = 'DESCONSIDERAR'
        
    df_atual.sort_values(by='ENTRADA',inplace=True)
    return df_atual


def virando_datas(df):
    df.loc[df['STATUS']=='ATIVO','ENTRADA'] = df['FIM DA VIG']
    df.loc[df['STATUS']=='ATIVO','FIM DA VIG'] = df['ENTRADA'] + pd.DateOffset(days=30)

    return df