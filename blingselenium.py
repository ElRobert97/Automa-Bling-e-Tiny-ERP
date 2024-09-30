#!/usr/bin/env python
# coding: utf-8

# In[ ]:

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementNotVisibleException,TimeoutException, ElementNotInteractableException, StaleElementReferenceException      
import time 
import pandas as pd


filtro_status_pedidos = '//*[@id="search-tag"]/span[1]/span[2]/span[2]/i'
texto_venda = '//*[@id="datatable"]/table/tbody/tr/td[4]/span[2]'
texto_valor = '//*[@id="datatable"]/table/tbody/tr/td[5]'
btn_pesq_receber = 'pesquisa-mini'
btn_tres_pont_receber = '/html/body/div[6]/div[3]/div[3]/div[2]/table/tbody/tr/td[10]/div/button/i'
btn_baixar_total = '/html/body/div[6]/div[3]/div[3]/div[2]/table/tbody/tr/td[10]/div/ul/li[1]/a/span[2]'
campo_calendario = '/html/body/div[17]/div[2]/div[2]/div[2]/div/div[3]/div/input'


def iniciar(driver, login ,senha, tela_contas_receber):  
    #Aguardando a pagina de login
    while True:
            try:
                    login_bling = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="username"]'))
                    )
                    break
            except TimeoutException:
                    driver.refresh()
                    time.sleep(1)
                    
    senha_bling = driver.find_element(By.XPATH, '//*[@id="login"]/div/div[1]/div/div[2]/div/input')
    botao_entrar = driver.find_element(By.XPATH, '//*[@id="login"]/div/div[1]/div/button[1]')     
    login_bling.send_keys(login)
    senha_bling.send_keys(senha)
    botao_entrar.click()
    element = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, filtro_status_pedidos )))
                                        
    element.click()

    #Abrir nova janela
    driver.execute_script("window.open('');")

    #Trocando para o contas a receber
    driver.switch_to.window(driver.window_handles[1])
    driver.get(tela_contas_receber) 
    driver.switch_to.window(driver.window_handles[0])

def tela_espera(driver):
    while True:
            try:
                elemento = WebDriverWait(driver, 1).until(
                     EC.visibility_of_element_located((By.XPATH, '//*[@id="wait"]/div'))
                     )
            except (TimeoutException, ElementNotVisibleException):
                    break

def conciliacao(planilha, caixa_bling, coluna_data, coluna_pedidos, inicial, driver, nao_encontrados):
    try:
        situacao_pedidos = driver.find_element(By.XPATH, '/html/body/div[6]/div[7]/div[2]/div[1]/div[2]/span[1]/span[2]/span[2]/i')
        situacao_pedidos.click()
    except:
        pass    
    calendario = driver.find_element(By.ID, 'dtButton')
    calendario.click()
    data_ini = driver.find_element(By.XPATH, '/html/body/div[6]/div[7]/div[2]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/input')
    tela_espera(driver)
    actions = ActionChains(driver)
    data_ini.click()
    #Ctrl + A para selecionar tudo
    actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
    #Delete para apagar o conteúdo
    actions.send_keys(Keys.DELETE).perform()
    time.sleep(1)
    #Enviando a Data
    data_ini.send_keys('01/01/2024')
    btn_filtrar_por_vendas = driver.find_element(By.XPATH, '//*[@id="dialog-picker"]/div[3]/div[2]/button')
    btn_filtrar_por_vendas.click()
    
    
    for i, row in planilha.iterrows():
        data = row[coluna_data]
        pedido = row[coluna_pedidos]
         
        
        #Pesquisando o pedido de venda
        while True:
            try:   
                btn_pesq_vendas = driver.find_element(By.ID, 'psqNumeroPedidoDaLojaVirtual')
                btn_pesq_vendas.clear()
                btn_pesq_vendas.send_keys(str(inicial + pedido))
                btn_pesq_vendas.send_keys(Keys.ENTER)
                break
            except ElementNotInteractableException:
                btn_filtro = driver.find_element(By.ID , 'open-filter')
                btn_filtro.click()     
        #Pegando o nome do cliente
        tela_espera(driver)
       
        btn_tres_pontinhos_vendas = driver.find_element(By.XPATH, '/html/body/div[6]/div[7]/div[2]/div[2]/table/tbody/tr[1]/td[7]/div/button')
        btn_tres_pontinhos_vendas.click()
        botao_lancar_contas = driver.find_element(By.XPATH, '/html/body/div[6]/div[7]/div[2]/div[2]/table/tbody/tr/td[7]/div/ul/li[3]/a/span[2]')
        tela_espera(driver)
        if botao_lancar_contas.text == 'Lançar contas':
            botao_lancar_contas.click()
       
        venda = driver.find_element(By.XPATH ,texto_venda).text
        valor = driver.find_element(By.XPATH, texto_valor).text
        
        driver.switch_to.window(driver.window_handles[1])
        tela_espera(driver)
        #Pesquisando a contas a receber do cliente
        while True:
            try:
                element = WebDriverWait(driver, 60).until(
                        EC.visibility_of_element_located((By.ID,btn_pesq_receber))
                    )
                break
            except StaleElementReferenceException:
                btn_filtro.click()
                
        element.clear()
        element.send_keys(venda)
        pesq_valor_receb = driver.find_element(By.ID, 'valorPsq')
        pesq_valor_receb.clear()
        pesq_valor_receb.send_keys(valor)
        element.send_keys(Keys.ENTER)
        tela_espera(driver)
                
            
        #Aguardando os 3 pontinhos da conta a receber para baixar
        try:
            element = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[6]/div[3]/div[3]/div[2]/div/div/h3'))
                )
            if element.text == 'Nenhum resultado encontrado.':
                nao_encontrados['Pedidos'].append(pedido)
                driver.switch_to.window(driver.window_handles[0]) 
                continue
        except:
            btn_3_pontos = driver.find_element(By.XPATH, btn_tres_pont_receber)
            btn_3_pontos.click()

        #Aguando o elemento de baixar a conta totalmente
        element = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, btn_baixar_total))
                                            )
        element.click()

        tela_espera(driver)
        #Limpando a Data e inserindo a data de recebimento
        element = WebDriverWait(driver, 60).until(
                EC.visibility_of_element_located((By.XPATH, campo_calendario))
                                            )
        actions = ActionChains(driver)
        element.click()
        #Ctrl + A para selecionar tudo
        actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        #Delete para apagar o conteúdo
        actions.send_keys(Keys.DELETE).perform()
        time.sleep(1)
        #Enviando a Data
        element.send_keys(str(data))

        #Preenchendo a categoria
        categoria = driver.find_element(By.NAME, 'portador-bill')
        categoria.send_keys(caixa_bling)
        categoria.send_keys(Keys.ENTER)
        #Baixando a conta
        btn_baixar_ok = driver.find_element(By.ID, 'btnBaixarBordero')
        btn_baixar_ok.click()
        #Retornando a tela de pedidos de venda
        driver.switch_to.window(driver.window_handles[0]) 
        
    df = pd.DataFrame(nao_encontrados)
    df.to_excel('Não encontrados.xlsx', index = False)


# In[64]:





