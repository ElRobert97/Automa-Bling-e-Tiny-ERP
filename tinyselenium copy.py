import sys
import os

# Adicione o diretório raiz ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from arquivos_principais.tinyselenium import iniciar, processar_pedidos
from arquivos_principais.ajustar_planilhas import tratar_magalu
from selenium import webdriver


login = r'fiscal@kibunitinho'
senha = r'@Contabilidade2024'
categoria = r'Clientes - Revenda de Mercadoria'

caixa_magalu = 'Magalu'
caminho_magalu = r'Kibunitinho\marketplaces\Magalu.csv'
planilha_magalu = tratar_magalu(caminho_magalu)
driver = webdriver.Chrome() 
iniciar(login,senha, driver=driver)
processar_pedidos(planilha = planilha_magalu, caixa_tiny = caixa_magalu, categoria = categoria , col_pedido = 'pedido', col_data = 'data de pagamento', inicial = 'LU-', driver = driver)
