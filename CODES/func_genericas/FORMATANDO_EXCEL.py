import pandas as pd
import string
import os

dict_columns_format = {

     'BOX': {
            'TOTAL_TYPE':'total_string',
            'TOTAL_VALUE':'TOTAIS'
            }, 
     
     'NOME': {
            'TOTAL_TYPE':'total_function',
            'TOTAL_VALUE':'count'
            }, 
     
     'INCÊNDIO': {
            'TOTAL_TYPE':'total_function',
            'TOTAL_VALUE':'sum'
            }, 

     'PRÊMIO': {
            'TOTAL_TYPE':'total_function',
            'TOTAL_VALUE':'sum'
            }, 
    
     'RECEBIMENTO JR': {
            'TOTAL_TYPE':'total_function',
            'TOTAL_VALUE':'count'
            }
    }


def main(path_saida,path_atual,df_atual,label_1,label_2,label_3,tipo='ATUAL'):

    # VALIDA SE O EXCEL NÃO ESTÁ VAZIO, PARA FORMATÁ-LO
    if df_atual.empty == False:
        
        # CRIANDO DIRETÓRIO CASO NÃO EXISTA
        dir_saida = '/'.join(path_saida.split('/')[:-1]) + '/'
        if not os.path.exists(dir_saida):
            os.makedirs(dir_saida)
        
        # LENDO BASE DE CLIENTES CANCELADOS
        if tipo == 'ATUAL':
            df_base_clientes_cancelados = ler_cancelados_atual(path_saida)
        elif tipo == 'VIRADA':
            df_base_clientes_cancelados = ler_cancelados_para_virada(path_atual)
        elif tipo == 'PRISMA':
            df_base_clientes_cancelados = None

        # CHAMANDO EXCEL WRITER
        writer = criando_excel_saida(path_saida)

        # CRIANDO SHEET ATIVOS
        sheet_ativos(writer,df_atual,label_1,label_2,label_3,tipo)
        sheet_cancelados(writer,df_atual,df_base_clientes_cancelados,tipo,label_1,label_2,label_3)

        while True:
            try:
                writer.save()
                break
            except:
                print("\n\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n\n\nERRO NA HORA DE SALVAR O EXCEL. FECHE A PREVISÃO OU FATURA DESTE STORAGE\n\n\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n\n")
                input('FECHE O EXCEL E APERTE ENTER')

    else:
        print('FORMATAÇÃO NÃO REALIZADA, POIS A BASE ESTÁ VAZIA.')


def ler_cancelados_atual(path_saida):
    try:
        df_base_clientes_cancelados = pd.read_excel(path_saida,skiprows=3,sheet_name='CANCELADOS')
    except:
        df_base_clientes_cancelados = None
    
    return df_base_clientes_cancelados

def ler_cancelados_para_virada(path_atual):
    try:
        if os.path.exists(path_atual):
            df_base_clientes_cancelados = pd.read_excel(path_atual,skiprows=3,sheet_name='CANCELADOS')
        else:
            df_base_clientes_cancelados = pd.read_excel(path_atual.replace('P_','F_'),skiprows=3,sheet_name='CANCELADOS')
            
    except:
        df_base_clientes_cancelados = None
    
    return df_base_clientes_cancelados
        

def sheet_cancelados(writer,df_atual,df_cancelados,tipo,label_1,label_2,label_3):
    
    # VALIDANDO CONJUNTO DE CANCELADOS ENTRE AS BASE DE ATVOS/JÁ CANCELADOS
    if tipo != 'PRISMA':
        try:
            df_base_clientes = df_cancelados.append(df_atual,ignore_index=True)
            df_base_clientes = df_base_clientes.loc[df_base_clientes['STATUS']=='CANCELADO'].copy()
        except:
            df_base_clientes = df_atual.loc[df_atual['STATUS']=='CANCELADO'].copy()
    else:
        df_base_clientes = df_atual.loc[df_atual['STATUS']=='CANCELADO'].copy()
    
    # DROP NAS LINHAS COM BOX EM BRANCO
    df_base_clientes.dropna(axis=0,subset=['BOX'],inplace=True)

    # EXPORTANDO DATAFRAME PARA XLSX
    df_base_clientes.to_excel(writer, sheet_name='CANCELADOS',index=False, startrow=3)
    
    # DEFININDO FORMATAÇÕES DA TABELA
    column_settings = [{'header': column,} for column in df_base_clientes.columns]
    
    # INSERINDO TABELA NA SHEET
    (max_row, max_col) = df_base_clientes.shape
    if max_row == 0:
        max_row = 1
    worksheet = writer.sheets['CANCELADOS']
    worksheet.add_table(3, 0, max_row+3, max_col - 1,{'columns': column_settings,'style': 'Table Style Light 16'})

    # DEFININDO COLUNAS RECEBIMENTO E ÚLTMA COBERTURA
    all_letters = list(string.ascii_uppercase)
    col_recebimento = all_letters[max_col-1]
    col_last_cobertura = all_letters[max_col-2]

    # CRIANDO WORKSHEET E EDITANDO COLUNAS
    workbook  = writer.book
    worksheet = writer.sheets['CANCELADOS']
    worksheet.set_zoom(75)
    worksheet.set_column('B:B', 18)
    worksheet.set_column('C:C', 38)
    worksheet.set_column('D:D', 20)
    worksheet.set_column(col_recebimento+':'+col_recebimento, 20)
   
    # CRIANDO FORMATOS DE CÉLULAS
    money = workbook.add_format({'num_format': 44, 'font_color': 'black'})
    alignment = workbook.add_format({'align':'center','valign':'vcenter','font_color':'black'})
    alignment2 = workbook.add_format({'align':'left','font_color':'black'})
    format_jr = workbook.add_format({'bg_color':'yellow','num_format':'dd/mm/yyyy'})
    
    # FORMATANDO COLUNAS
    worksheet.set_column('A:' + col_recebimento, None, cell_format=alignment)
    worksheet.set_column('C:C', 38, cell_format=alignment2)
    worksheet.set_row(3,cell_format=alignment)

    # VARIAÇÕES DE FORMATAÇÕES ENTRE SEGURADORAS
    if 'ENTRADA' in df_base_clientes.columns:
        worksheet.set_column('G:' + col_last_cobertura, 20, cell_format=money)
        worksheet.set_column('E:F', 15)
    else:
        worksheet.set_column('E:' + col_last_cobertura, 20, cell_format=money)

    # SETANDO RANGE PARA FORMATAÇÃO DA COLUNA DE RECEBIMENTO
    range_recebimento_jr = col_recebimento + '5:' + col_recebimento + str(max_row+4)
    worksheet.conditional_format(range_recebimento_jr, {'type': 'no_errors','format': format_jr})
    
    # CRIANDO ESTILOS PARA AS FORMATAÇÕES CONDICIONAIS
    formato_verde = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100','align':'center'})
    formato_amarelo = workbook.add_format({'bg_color': '#ffeb9c','font_color': '#9c5700','align':'center'})
    formato_vermelho = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'})

    # DEFININDO RANGES PARA FORMATAÇÕES CONDICIONAIS
    range_status = 'B5:'+'B'+str(max_row+4)

    # CRIANDO FORMATAÇÕES CONDICIONAIS NOS RANGES
    worksheet.conditional_format(range_status, {'type': 'cell', 'criteria':'==', 'value':'"ATIVO"','format':formato_verde})
    worksheet.conditional_format(range_status, {'type': 'cell', 'criteria':'==', 'value':'"CANCELADO"','format':formato_vermelho})
    worksheet.conditional_format(range_status, {'type': 'cell', 'criteria':'==', 'value':'"ÚLTIMA COBRANÇA"','format':formato_amarelo})
    
    # CRIANDO FORMATAÇÃO DOS LABELS
    format_cabecalho = workbook.add_format({'font_color': 'black','align':'center','valign':'vcenter','font_size':20,'bold':True,'border':1})
    
    # DEFININDO LABELS
    worksheet.merge_range("A1:"+col_recebimento+"1", label_1, format_cabecalho)
    worksheet.merge_range("A2:"+col_recebimento+"2", label_2, format_cabecalho)
    worksheet.merge_range("A3:"+col_recebimento+"3", label_3, format_cabecalho)


def sheet_ativos(writer,df_base_clientes,label_1,label_2,label_3,tipo):
    
    # LENDO APENAS ATIVOS
    if tipo == 'VIRADA':
        df_base_clientes_ativos = df_base_clientes.loc[df_base_clientes['STATUS']=='ATIVO'].copy()
    else:
        df_base_clientes_ativos = df_base_clientes.loc[~(df_base_clientes['STATUS'].isin(['CANCELADO','DESCONSIDERAR']))].copy()
    
    # EXPORTANDO DATAFRAME PARA XLSX
    df_base_clientes_ativos.to_excel(writer, sheet_name='ATIVOS',index=False, startrow=3)
    
    # DEFININDO FORMATAÇÕES DA TABELA
    columns = df_base_clientes_ativos.columns
    column_settings = definindo_dict_columns(columns,dict_columns_format)
    
    # INSERINDO TABELA NA SHEET
    (max_row, max_col) = df_base_clientes_ativos.shape
    if max_row == 0:
        max_row = 1
    worksheet = writer.sheets['ATIVOS']
    worksheet.add_table(3, 0, max_row+4, max_col - 1,{'columns': column_settings,'style': 'Table Style Light 16','total_row':True})

    # DEFININDO COLUNAS RECEBIMENTO E ÚLTMA COBERTURA
    all_letters = list(string.ascii_uppercase)
    col_recebimento = all_letters[max_col-1]
    col_last_cobertura = all_letters[max_col-2]

    # CRIANDO WORKSHEET E EDITANDO COLUNAS
    workbook  = writer.book
    worksheet = writer.sheets['ATIVOS']
    worksheet.set_zoom(75)
    worksheet.set_column('B:B', 18)
    worksheet.set_column('C:C', 38)
    worksheet.set_column('D:D', 20)
    worksheet.set_column(col_recebimento+':'+col_recebimento, 20)
   
    # CRIANDO FORMATOS DE CÉLULAS
    money = workbook.add_format({'num_format': 44, 'font_color': 'black'})
    alignment = workbook.add_format({'align':'center','valign':'vcenter','font_color':'black'})
    alignment2 = workbook.add_format({'align':'left','font_color':'black'})
    format_jr = workbook.add_format({'bg_color':'yellow','num_format':'dd/mm/yyyy'})
    
    # FORMATANDO COLUNAS
    worksheet.set_column('A:' + col_recebimento, None, cell_format=alignment)
    worksheet.set_column('C:C', 38, cell_format=alignment2)
    worksheet.set_row(3,cell_format=alignment)

    # VARIAÇÕES DE FORMATAÇÕES ENTRE SEGURADORAS
    if 'ENTRADA' in df_base_clientes.columns:
        worksheet.set_column('G:' + col_last_cobertura, 20, cell_format=money)
        worksheet.set_column('E:F', 15)
    else:
        worksheet.set_column('E:' + col_last_cobertura, 20, cell_format=money)

    # SETANDO RANGE PARA FORMATAÇÃO DA COLUNA DE RECEBIMENTO
    range_recebimento_jr = col_recebimento + '5:' + col_recebimento + str(max_row+4)
    worksheet.conditional_format(range_recebimento_jr, {'type': 'no_errors','format': format_jr})
    
    # CRIANDO ESTILOS PARA AS FORMATAÇÕES CONDICIONAIS
    formato_verde = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100','align':'center'})
    formato_amarelo = workbook.add_format({'bg_color': '#ffeb9c','font_color': '#9c5700','align':'center'})
    formato_vermelho = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'})

    # DEFININDO RANGES PARA FORMATAÇÕES CONDICIONAIS
    range_status = 'B5:'+'B'+str(max_row+4)
    range_box = 'A5:'+'A'+str(max_row+4)

    # CRIANDO FORMATAÇÕES CONDICIONAIS NOS RANGES
    worksheet.conditional_format(range_box, {'type': 'duplicate','format':formato_vermelho})
    worksheet.conditional_format(range_status, {'type': 'cell', 'criteria':'==', 'value':'"ATIVO"','format':formato_verde})
    worksheet.conditional_format(range_status, {'type': 'cell', 'criteria':'==', 'value':'"CANCELADO"','format':formato_vermelho})
    worksheet.conditional_format(range_status, {'type': 'cell', 'criteria':'==', 'value':'"ÚLTIMA COBRANÇA"','format':formato_amarelo})
    
    # CRIANDO FORMATAÇÃO DOS LABELS
    format_cabecalho = workbook.add_format({'font_color': 'black','align':'center','valign':'vcenter','font_size':20,'bold':True,'border':1})
    
    # DEFININDO LABELS
    worksheet.merge_range("A1:"+col_recebimento+"1", label_1, format_cabecalho)
    worksheet.merge_range("A2:"+col_recebimento+"2", label_2, format_cabecalho)
    worksheet.merge_range("A3:"+col_recebimento+"3", label_3, format_cabecalho)


def criando_excel_saida(path_saida):
    writer =  pd.ExcelWriter(path_saida, engine='xlsxwriter',datetime_format="dd/mm/yyyy") 
    return writer


def definindo_dict_columns(list_of_columns,dict_columns_format):
    config_columns = []
    for column in list_of_columns:
        dict_temp = {}
        x = dict_columns_format.get(column)
        if x != None:
            type_ = x.get('TOTAL_TYPE')
            value = x.get('TOTAL_VALUE')
        else:
            type_ = 'total_string'
            value = ''
        
        dict_temp['header'] = column
        dict_temp[type_] = value
        config_columns.append(dict_temp)
    
    return config_columns
