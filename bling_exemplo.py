import sys
import os
# Adicione o diretório raiz ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Agora, importe os módulos
from arquivos_principais.blingselenium import conciliacao, iniciar
from arquivos_principais.ajustar_planilhas import tratar_amazon, tratar_magalu, tratar_shein, tratar_ML
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def driver_generico():
    driver = webdriver.Chrome(options = options) 
    return driver

options = Options()
options.add_argument("--disable-blink-features=AutomationControlled")  # Evitar detecção do Selenium
# options.add_argument(r"user-data-dir=C:\Users\Contler Elias\OneDrive - AGORA DEU LUCRO EDUCACIONAL LTDA\Área de Trabalho\Perfil_Navegador") # Diretorio do Perfil
# options.add_argument(r"profile-directory=Profile 10") # Perfil selecionado

# Login ERP 
login = ''
senha = ''
categoria = ''

dia = '10'
mes = '9'
ano = '2024'

data_final = (ano,mes,dia)


caminho_amazon= 'Albha\marketplaces\Amazon.csv'
planilha_amazon = tratar_amazon(caminho_amazon)
caixa_amazon = 'Amazon Pay'


caminho_magalu= 'Albha\marketplaces\Magalu.csv'
planilha_magalu= tratar_magalu(caminho_magalu)
caixa_magalu = 'Magalu Pay'

caminho_shein = 'Albha\marketplaces\Shein.xlsx'
planilha_shein = tratar_shein(caminho_shein)
caixa_shein= 'Shein'

# caminho_ml = 'Albha\marketplaces\ML_desdejunho.csv'
# planilha_ml = tratar_ML(caminho_ml, data_final)
# caixa_ml = 'Mercado Livre'

driver = driver_generico()
def conciliar(driver):
    iniciar(driver, login ,senha)
    #conciliacao(planilha=planilha_ml,caixa_bling=caixa_ml,coluna_data='Data de liberação do dinheiro (date_released)', coluna_pedidos='Número da venda no Mercado Livre (order_id)',inicial='', driver=driver)   
    conciliacao(planilha=planilha_magalu,caixa_bling=caixa_magalu,coluna_data='data de pagamento', coluna_pedidos='pedido',inicial='LU-',driver=driver, categoria_vendas= categoria)
    conciliacao(planilha=planilha_amazon,caixa_bling=caixa_amazon,coluna_data='data/hora', coluna_pedidos='id do pedido',inicial='',driver=driver, categoria_vendas= categoria)
    conciliacao(planilha=planilha_shein,caixa_bling=caixa_shein,coluna_data='Data de conclusão da liquidação', coluna_pedidos='Número do pedido correspondente',inicial='', driver=driver, categoria_vendas= categoria)
    

conciliar(driver)


