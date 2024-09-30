import sys
import os

# Adicione o diretório raiz ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Agora, importe os módulos
from arquivos_principais.blingselenium import conciliacao, tela_espera, iniciar
from arquivos_principais.ajustar_planilhas import tratar_shopee
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument("--disable-blink-features=AutomationControlled")  # Evitar detecção do Selenium
options.add_argument(r"user-data-dir=C:\Users\Contler Elias\AppData\Local\Google\Chrome\User Data") # Diretorio do Perfil
options.add_argument(r"profile-directory=Profile 60") # Perfil selecionado

tela_pedidos_vendas = 'https://www.bling.com.br/vendas.php#list'
tela_contas_receber = 'https://www.bling.com.br/contas.receber.php'

nao_encontrados = {'Pedidos':[]}
caixa_bling = 'Shopee'
login = 'agoradeulucro@maiconsuporti'
senha = 'Agoradeulucro24@'

caminho = 'Albha\marketplaces'
caminho_amazon= caminho + ''
planilha_amazon = tratar_shopee(caminho_amazon)

caminho = 'Albha\marketplaces'
caminho_shein= caminho + ''
planilha_shein = tratar_shopee(caminho_shein)

caminho = 'Albha\marketplaces'
caminho_magalu= caminho + ''
planilha_magalu= tratar_shopee(caminho_magalu)

driver = webdriver.Chrome(options = options) 
driver.get(tela_pedidos_vendas)

tela_espera(driver)
iniciar(driver, login ,senha, tela_contas_receber)

conciliacao(planilha=planilha_amazon,caixa_bling=caixa_bling,coluna_data='Data de conclusão do pagamento', coluna_pedidos='ID do pedido',inicial='', driver=driver, nao_encontrados=nao_encontrados)
