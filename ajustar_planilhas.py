import pandas as pd
from datetime import datetime
import re
import locale
from datetime import date
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

#Tratando Planilha Olist
def transf_hora_ml(data_str):
        # Converte a string para datetime    
        data_obj = datetime.strptime(data_str, "%d/%m/%Y %H:%M:%S")
        data_nova = datetime.strftime(data_obj, '%d/%m/%Y')
        return data_nova

def tratar_ML(planilha, data_final):
        ano, mes, dia = data_final
        data_limite = date(int(ano), int(mes), int(dia))
        planilha_ml = pd.read_csv(planilha, sep = ';')
        planilha_ml = planilha_ml[['Data de liberação do dinheiro (date_released)', 'Status da operação (status)', 'Número da venda no Mercado Livre (order_id)']]
        planilha_filtrada = planilha_ml.loc[planilha_ml['Status da operação (status)'] == 'approved' ,:].copy()
        planilha_filtrada['Data de liberação do dinheiro (date_released)'] = pd.to_datetime(planilha_filtrada['Data de liberação do dinheiro (date_released)'], format ='%d/%m/%Y %H:%M:%S' , errors = 'coerce')
        planilha_filtrada['Data de liberação do dinheiro (date_released)'] = planilha_filtrada['Data de liberação do dinheiro (date_released)'].dt.date
        planilha_filtrada['Número da venda no Mercado Livre (order_id)'] = planilha_filtrada['Número da venda no Mercado Livre (order_id)'].astype(str)
        planilha_filtrada_n = planilha_filtrada.loc[planilha_filtrada['Data de liberação do dinheiro (date_released)'] <= data_limite ,:].copy()
        planilha_filtrada_n['Data de liberação do dinheiro (date_released)'] = pd.to_datetime(planilha_filtrada['Data de liberação do dinheiro (date_released)'], format ='%d/%m/%Y').dt.strftime('%d/%m/%Y')
        return planilha_filtrada_n

def transf_data_olist(data):
    
    nova_data = datetime.strptime(data, "%d/%m/%Y %Hh%M").date()
    data_formatada = datetime.strftime(nova_data, '%d/%m/%Y')
    
    return data_formatada

def tratar_olist(planilha):
    planilha_olist = pd.read_excel(planilha, skiprows = 2, sheet_name= 'Resumo Mensal (previsto)')
    planilha_olist['data da transação'] = planilha_olist['data da transação'].apply(transf_data_olist)
    planilha_olist_filtrada = planilha_olist.loc[planilha_olist['tipo da transação'] == 'repasse', :].copy()
    planilha_olist_filtrada = planilha_olist[['ciclo de repasse','tipo da transação','pedido']]
    planilha_olist_filtrada['pedido'] = planilha_olist_filtrada['pedido'].astype(str)
    return planilha_olist_filtrada


#Tratando Planilha da Magalu
def tratar_magalu(planilha):
    planilha_magalu = pd.read_excel(planilha)
    planilha_magalu['Data da transação'] = pd.to_datetime(planilha_magalu['Data da transação'], format ='%d/%m/%Y %H:%M').dt.strftime('%d/%m/%Y')
    planilha_magalu = planilha_magalu.dropna(subset = ['ID do pedido Seller'])
    return planilha_magalu

#Tratando Planilha da Amazon
def transf_hora_amazon(data_str):
        if len(data_str) == 33:
            data_str = data_str[:18]
        else:
            data_str = data_str[:17]
        data_str = data_str.replace('.', '')
        sub_data = re.sub(r'(\b\d{1}\b)( de \w+ de \d{4})', r'0\1\2', data_str)
        sub_data = sub_data.strip()
        
        # Converte a string para datetime    
        data_obj = datetime.strptime(sub_data, "%d de %b de %Y").date()
        data_nova = datetime.strftime(data_obj, '%d/%m/%Y')
        return data_nova
    

def tratar_amazon(planilha_amazon):
    amazon = pd.read_csv(planilha_amazon, skiprows = 7)
    amazon_filtrada = amazon[['data/hora', 'tipo', 'id do pedido']].copy()
    amazon_filtrada = amazon_filtrada.loc[amazon['tipo'] == 'Pedido', :]
    amazon_filtrada['data/hora'] = amazon_filtrada['data/hora'].apply(transf_hora_amazon)
    return amazon_filtrada
    
    
#Gerar planilha de não encontrados
def plan_nao_encontrados(nao_encontrados):
    df = pd.DataFrame(nao_encontrados)
    df.to_excel('Pedidos_nao_encontrados.xlsx', index = False)
    return df


def transf_data_shopee(data):
    data = datetime.strptime(data, '%Y-%m-%d')
    data_nova = datetime.strftime(data, '%d/%m/%Y')
    return data_nova

def tratar_shopee(planilha_shopee):
    planilha = pd.read_excel(planilha_shopee, sheet_name= 'Renda', skiprows= 2)
    planilha = planilha.loc[planilha['Ver'] == 'Order', :]
    planilha['Data de conclusão do pagamento'] = planilha['Data de conclusão do pagamento'].apply(transf_data_shopee)
    return planilha

def transf_data_shein(data): 
    data_nova = str(data)
    data_convertida = datetime.strptime(data_nova, "%d %B %Y")
    nova_data = datetime.strftime(data_convertida, '%d/%m/%Y')
    return nova_data

def tratar_shein(planilha_shein):
    planilha = pd.read_excel(planilha_shein)
    planilha = planilha.drop([0])
    planilha_filtrada = planilha[['Número do pedido correspondente', 'Tipo de fatura', 'Data de conclusão da liquidação']].copy()
    planilha_filtrada['Data de conclusão da liquidação'] = planilha_filtrada['Data de conclusão da liquidação'].apply(transf_data_shein)
    planilha_filtrada = planilha_filtrada.loc[planilha['Tipo de fatura'] == 'Receita do pedido', :]
    
    return  planilha_filtrada

def tratar_via(planilha_via):
    planilha = pd.read_csv(planilha_via, sep = ';')
    planilha_filtrada = planilha.loc[planilha['Tipo da Transação'] == 'VENDA', :]
    planilha_filtrada['Número Pedido'] = planilha_filtrada['Número Pedido'].astype(str)
    return planilha_filtrada


def tratar_nuvemshop(planilha_nuvemshop):
    planilha = pd.read_csv(planilha_nuvemshop, encoding = 'latin1', sep = ';')
    planilha = planilha[['Status do Pedido', 'Status do Pagamento','Data','Identificador do pedido', 'Forma de Pagamento']]
    planilha['Identificador do pedido'] = planilha['Identificador do pedido'].astype(str)
    planilha['Identificador do pedido'] = planilha['Identificador do pedido'].replace('.0','').apply(lambda x: x.replace('.0',''))
    planilha_filtrada = planilha.loc[(planilha['Status do Pagamento'] == 'Confirmado') & (planilha['Forma de Pagamento'] == 'Mercado Pago'),:]
    return planilha_filtrada


def transf_data_kabum(data_str):
        # Converte a string para datetime 
        data_obj = datetime.strptime(str(data_str), "%d/%m/%Y - %H:%M:%S")
        data_nova = datetime.strftime(data_obj, '%d/%m/%Y')
        return data_nova


def tratar_kabum(caminho):
        # Converte a string para datetime    
        planilha = pd.read_csv(caminho, sep = ';', decimal = '.')
        planilha = planilha.dropna(subset='Data recebida')
        planilha['Data recebida'] = planilha['Data recebida'].apply(transf_data_kabum)
        planilha['N° do pedido'] = planilha['N° do pedido'].astype(str)
        planilha_filtrada = planilha.loc[planilha['Tipo'] == 'Valor do pedido', :]
        return planilha_filtrada