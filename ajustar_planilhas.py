import pandas as pd
from datetime import datetime
import re
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
#Tratando Planilha Olist

def transf_data_olist(data):
    
    nova_data = datetime.strptime(data, "%d/%m/%Y %Hh%M").date()
    data_formatada = datetime.strftime(nova_data, '%d/%m/%Y')
    
    return data_formatada

def tratar_olist(planilha):
    planilha_olist = pd.read_excel('Olist_Veg.xlsx', skiprows = 2, sheet_name= 'Resumo Mensal (previsto)')
    planilha_olist['data da transação'] = planilha_olist['data da transação'].apply(transf_data_olist)
    planilha_olist_filtrada = planilha_olist.loc[planilha_olist['tipo da transação'] == 'repasse', :]
    planilha_olist_filtrada = planilha_olist[['ciclo de repasse','tipo da transação','pedido']]
    return planilha_olist_filtrada


#Tratando Planilha da Magalu
def tratar_magalu(planilha):
    planilha_magalu = pd.read_csv(planilha)
    planilha_magalu = planilha_magalu.loc[planilha_magalu['tipo de operação'] == 'transação' , :]
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
    amazon_filtrada = amazon[['data/hora', 'tipo', 'id do pedido']]
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
    planilha = pd.read_excel(planilha_shopee, sheet_name= 'Income', skiprows= 5)
    planilha['Data de conclusão do pagamento'] = planilha['Data de conclusão do pagamento'].apply(transf_data_shopee)
    return planilha