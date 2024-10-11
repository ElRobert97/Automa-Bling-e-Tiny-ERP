#!/usr/bin/env python
# coding: utf-8

# In[ ]:

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException,ElementNotVisibleException,TimeoutException, ElementNotInteractableException, StaleElementReferenceException      
import time 
import pandas as pd


filtro_status_pedidos = '//*[@id="search-tag"]/span[1]/span[2]/span[2]/i'
texto_venda = '//*[@id="datatable"]/table/tbody/tr/td[4]/span[2]'
texto_valor = '//*[@id="datatable"]/table/tbody/tr/td[5]'
btn_pesq_receber = 'pesquisa-mini'
btn_tres_pont_receber = 'dropdown-toggle'
btn_baixar_total = '//td[10]/div/ul/li[1]/a/span[2]'
campo_calendario = '//*[starts-with(@id, "dp")]'

def clicar_btn_tres_pontinhos(driver):
    xpaths = [
        #'//*[contains(@class, "dropdown-toggle")]/i',  # Primeira opção
        #'//td[7]/div/button/i',                        # Segunda opção
        #'/html/body/div[6]/div[7]/div[2]/div[2]/table/tbody/tr/td[7]/div/button/i',  # Terceira opção
        '//*[@id="datatable"]/table/tbody/tr/td[7]/div/button',  # Quarta opção
        #'//*[@id="datatable"]/table/tbody/tr/td[7]/div/ul'       # Quinta opção
    ]
    for xpath in xpaths:
        try:
            btn_tres_pontinhos_vendas = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            )
            btn_tres_pontinhos_vendas.send_keys(Keys.ENTER)
            break  # Sai do loop se o botão for clicado com sucesso
        except:
            continue  # Tenta o próximo xpath se o atual falhar

def clicar_btn_tres_pontinhos_nf(driver):
    
    xpaths = ['//*[contains(@id)]/td[8]/div/ul',
              '//*[contains(@id)]/td[8]/div',
              '//*[contains(@class, "dropdown-toggle")]/i',
              '//*[contains(@id)]/td[8]'
    ]
    for xpath in xpaths:
        try:
            botao_tres_pontinhos = driver.find_element(By.XPATH, xpath )
            botao_tres_pontinhos.send_keys(Keys.ENTER)
            break  # Sai do loop se o botão for clicado com sucesso
        except:
            continue  # Tenta o próximo xpath se o atual falhar


def iniciar(driver, login ,senha):  
    #Abrindo os pedidos de venda
    driver.get('https://www.bling.com.br/vendas.php#list')
    driver.maximize_window()
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
    login_bling.clear()
    login_bling.send_keys(login)
    senha_bling.clear()
    senha_bling.send_keys(senha)
    botao_entrar.click()

    element = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, filtro_status_pedidos )))
                                        
    element.click()
  
    try:
        situacao_pedidos = driver.find_element(By.XPATH, '/html/body/div[6]/div[7]/div[2]/div[1]/div[2]/span[1]/span[2]/span[2]/i')
        situacao_pedidos.click()
    except:
        time.sleep(1)
              
    calendario = driver.find_element(By.ID, 'dtButton')
    calendario.click()
    data_ini = driver.find_element(By.XPATH, '/html/body/div[6]/div[7]/div[2]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/input')
    
    actions = ActionChains(driver)
    data_ini.click()
    #Ctrl + A para selecionar tudo
    actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
    #Delete para apagar o conteúdo
    actions.send_keys(Keys.DELETE).perform()
    time.sleep(1)
    #Enviando a Data
    data_ini.send_keys('01/10/2023')
    data_ini.send_keys(Keys.ENTER)
    # btn_filtrar_por_vendas = driver.find_element(By.XPATH, '//*[@id="dialog-picker"]/div[3]/div[2]/button')
    # btn_filtrar_por_vendas.click()

    #Abrir nova janela
    driver.execute_script("window.open('');")
    #Trocando para o contas a receber
    driver.switch_to.window(driver.window_handles[1])
    driver.get('https://www.bling.com.br/contas.receber.php')
    
    calendario = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, 'dtButton'))
        )
    calendario.click()
    data_ini = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, 'data-ini'))
    )
    data_fim = WebDriverWait(driver,15).until(
        EC.visibility_of_element_located((By.ID, 'data-fim'))
    )
    data_fim.click()
    data_fim.send_keys('31/12/2024')
    actions = ActionChains(driver)
    data_ini.click()
    #Ctrl + A para selecionar tudo
    actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
    #Delete para apagar o conteúdo
    actions.send_keys(Keys.DELETE).perform()
    time.sleep(1)
    #Enviando a Data
    data_ini.send_keys('01/01/2024')
    actions = ActionChains(driver)
    data_fim.click()
    #Ctrl + A para selecionar tudo
    actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
    #Delete para apagar o conteúdo
    actions.send_keys(Keys.DELETE).perform()
    data_fim.send_keys('31/12/2024')
    data_fim.send_keys(Keys.ENTER)
    # btn_filtrar_por_vendas = driver.find_element(By.XPATH, '//*[@id="dialog-picker"]/div[3]/div[2]/button')
    # btn_filtrar_por_vendas.click() 
    
    #Abrir nova janela para as notas fiscais
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[2])
    driver.get('https://www.bling.com.br/notas.fiscais.php#list')
    
    calendario = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, 'dtButton'))
        )
    calendario.click()
    data_ini = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, 'data-ini'))
    )
    
    actions = ActionChains(driver)
    data_ini.click()
    #Ctrl + A para selecionar tudo
    actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
    #Delete para apagar o conteúdo
    actions.send_keys(Keys.DELETE).perform()
    time.sleep(1)
    #Enviando a Data
    data_ini.send_keys('01/01/2024')
    data_fim = WebDriverWait(driver,15).until(
        EC.visibility_of_element_located((By.ID, 'data-fim'))
    )
    actions = ActionChains(driver)
    data_fim.click()
    #Ctrl + A para selecionar tudo
    actions.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
    #Delete para apagar o conteúdo
    actions.send_keys(Keys.DELETE).perform()
    data_fim.send_keys('31/12/2024')
    data_fim.send_keys(Keys.ENTER)
    # btn_filtrar_por_vendas = driver.find_element(By.XPATH, '//*[@id="dialog-picker"]/div[3]/div[2]/button')
    # btn_filtrar_por_vendas.click()
    
    #Voltando a tela de pedidos de vendas
    driver.switch_to.window(driver.window_handles[0])
           
def conciliacao(planilha, caixa_bling, coluna_data, coluna_pedidos, inicial, driver, categoria_vendas):
    nao_encontrados = {'Pedidos':[], 'Ocorrido':[]
                       }
    driver.switch_to.window(driver.window_handles[0])
    for i, row in planilha.iterrows():
        data = row[str(coluna_data)]
        pedido = row[str(coluna_pedidos)]
         
        pedido = inicial + pedido
        #Pesquisando o pedido de venda
        while True:
            try:   
                btn_pesq_vendas = WebDriverWait(driver,5).until(
                EC.visibility_of_element_located((By.ID, 'psqNumeroPedidoDaLojaVirtual'))
                )
                btn_pesq_vendas.clear()
                btn_pesq_vendas.send_keys(str(pedido))
                driver.find_element(By.ID, 'pesquisa-mini').send_keys(Keys.ENTER)
                break
            except (TimeoutException,NoSuchElementException,ElementNotInteractableException):
                try:
                    btn_filtro = WebDriverWait(driver,5).until(
                    EC.visibility_of_element_located((By.ID , 'open-filter')))
                    driver.execute_script("arguments[0].click();" ,btn_filtro)
                except (TimeoutException,NoSuchElementException,ElementNotInteractableException):
                    btn_filtro = driver.find_element(By.ID , '//*[@id="link-pesquisa"]/span')
                    btn_filtro.send_keys(Keys.ENTER)
        
        contador = 0
        while True:
            try:
                elemento = WebDriverWait(driver, 1).until(
             EC.visibility_of_element_located((By.XPATH, '//*[@id="wait"]/div'))
             )
                contador += 1 
            except (TimeoutException, ElementNotVisibleException):
                    break
        if contador == 60:
            driver.find_element(By.ID, 'pesquisa-mini').send_keys(Keys.ENTER)
            
        #Verificando se existe o pedido de venda
        try:         
            msg_naoencontrado = WebDriverWait(driver,2).until(
                    EC.visibility_of_element_located((By.XPATH , '//*[@id="datatable"]/div/div/h3'))
                )
            continue
        except(NoSuchElementException, TimeoutException):
            pass
        #Verificando se o pedido foi cancelado
        situacao_pedido = WebDriverWait(driver, 1).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="datatable"]/table/tbody/tr/td[6]/span[2]/span/div/span'))
        )
        if 'Cancelado' in situacao_pedido.text or 'Devolvido' in situacao_pedido.text:
            continue
        #Pegando o nome do cliente
        try:
             venda = WebDriverWait(driver,2).until(
                EC.visibility_of_element_located((By.XPATH ,texto_venda))
            )
             venda = venda.text
             valor = driver.find_element(By.XPATH, texto_valor).text  
        except (StaleElementReferenceException,NoSuchElementException, TimeoutException):
            motivo = 'Não encontrado o pedido de venda'
            nao_encontrados['Ocorrido'].append(motivo)
            nao_encontrados['Pedidos'].append(pedido)
            continue
        
        #Tentar encontrar o simbolo de conta lançada
        try:
            conta_lancada = driver.find_element(By.XPATH, '//*[contains(@class, "bling-c-green") or contains(@class, "bling-c-midgree")]')
            driver.switch_to.window(driver.window_handles[1])
        except NoSuchElementException:
            #Tentativa de lançar contas
            try: 
                #Vai tentar lançar contas pela tela de pedido de vendas
                tres_pontos = driver.find_element(By.CLASS_NAME, 'dropdown-toggle').click()
                selecao = WebDriverWait(driver, 5).until(
                            EC.visibility_of_all_elements_located((By.XPATH, '//ul[contains(@class, "dropdown-menu")]')) 
                        )
                encontrado = False
                for i in selecao:
                    texto = i.text
                    if 'Lançar contas' in texto:                  
                        botao_lancar_contas = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[text()='Lançar contas']"))
                        )
                        driver.execute_script("arguments[0].click();", botao_lancar_contas)
                        encontrado = True
                        driver.switch_to.window(driver.window_handles[1])
                        break
                #Altera para tela de nota fiscais para lançar conta
                if not encontrado:
                    driver.switch_to.window(driver.window_handles[2])
                    while True:
                        try:
                            pesquisa = WebDriverWait(driver, 5).until(
                                EC.visibility_of_element_located((By.ID, 'psqNumeroPedidoDaLojaVirtual'))
                            )
                            pesquisa.clear()
                            pesquisa.send_keys(pedido)
                            pesquisa.send_keys(Keys.ENTER)
                            break
                        except (TimeoutException,ElementNotInteractableException):
                            try:
                                btn_filtro = WebDriverWait(driver,5).until(
                                EC.visibility_of_element_located((By.ID , 'open-filter')))
                                btn_filtro.send_keys(Keys.ENTER)
                                break
                            except (TimeoutException,NoSuchElementException,ElementNotInteractableException):
                                    btn_filtro = driver.find_element(By.XPATH , '//*[@id="open-filter"]/span')
                                    btn_filtro.send_keys(Keys.ENTER)
                                
                    contador = 0
                    while True:
                        try:
                            elemento = WebDriverWait(driver, 1).until(
                            EC.visibility_of_element_located((By.XPATH, '//*[@id="wait"]/div'))
                                )
                            contador += 1 
                        except (TimeoutException, ElementNotVisibleException):
                                break
                    if contador == 60:
                        pesquisa.send_keys(Keys.ENTER)
                        
                        
                    #Acessando o menu para lançar contas
                    tres_pontos = WebDriverWait(driver,2).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'dropdown-toggle')))
                    tres_pontos.click()
                    selecao = WebDriverWait(driver, 2).until(
                        EC.visibility_of_all_elements_located((By.CLASS_NAME, 'dropdown-menu'))
                    )
                    for i in selecao:
                        texto = i.text
                        if 'Lançar contas' in texto:
                                botao_lancar_contas = WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, "//span[text()='Lançar contas']"))
                                )
                                driver.execute_script("arguments[0].click();", botao_lancar_contas)
                                driver.switch_to.window(driver.window_handles[1])
                                break
            except Exception as err:
                motivo = f'Não foi possivel lançar contas {err}'
                nao_encontrados['Ocorrido'].append(motivo)
                nao_encontrados['Pedidos'].append(pedido)
                driver.switch_to.window(driver.window_handles[1])
             
        
        #Verificando se o botao de alterar categoria esta presente
        try:
            confir_categoria = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, "//footer//button[contains(text(), ' Alterar todas as categorias ')]") )
            )
            confir_categoria.click()
        except:
            pass
        
        actions = ActionChains(driver)
        actions.send_keys(Keys.ESCAPE).perform()
        #Pesquisando a contas a receber do cliente
        element = WebDriverWait(driver, 60).until(
                        EC.visibility_of_element_located((By.ID,btn_pesq_receber))
                    )            
        element.clear()
        element.send_keys(venda)
        while True:
            try:
                pesq_valor_receb = WebDriverWait(driver, 3).until(
                                EC.visibility_of_element_located((By.ID, 'valorPsq'))
                            )
                pesq_valor_receb.clear()
                pesq_valor_receb.send_keys(valor)
                element.send_keys(Keys.ENTER)          
                break
            except (TimeoutException,StaleElementReferenceException):
                try:
                    btn_filtro = driver.find_element(By.XPATH , '//*[@id="open-filter"]')
                    btn_filtro.click()
                except (StaleElementReferenceException,NoSuchElementException,ElementNotInteractableException):
                    btn_filtro = driver.find_element(By.XPATH , '//*[@id="open-filter"]/span')
                    btn_filtro.send_keys(Keys.ENTER)
                            
        contador = 0
        while True:
            try:
                elemento = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="wait"]/div'))
                )
                contador += 1 
            except (TimeoutException, ElementNotVisibleException):
                    break
        if contador == 60:
            element.send_keys(Keys.ENTER)  
        #Aguardando os 3 pontinhos da conta a receber para baixar
        try:
            element = WebDriverWait(driver, 3).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, 'busca-nao-encontrada'))
                )                 
            driver.switch_to.window(driver.window_handles[0]) 
            continue
        except:
            btn_3_pontos = driver.find_element(By.CLASS_NAME, btn_tres_pont_receber)
            driver.execute_script("arguments[0].click();" ,btn_3_pontos)
                  
        
        #Aguando o elemento de baixar a conta totalmente
        element = WebDriverWait(driver, 60).until(
                EC.visibility_of_element_located((By.XPATH, btn_baixar_total))
                                            )
        element.click()
        while True:
            try:
                elemento = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="wait"]/div'))
                    )
                contador += 1 
            except (TimeoutException, ElementNotVisibleException):
                    break
        # Verificando se o botao de utilizar data de vencimento esta ativado
        
        while True:
            try:
                btn_vencimento = element = WebDriverWait(driver, 3).until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="slide-in-writeoff-grouped-bills"]/div[2]/div[1]/div/div[3]/div/div/label[2]'))
            )
                if btn_vencimento.text == 'Desativado':
                        pass
                        break
                else:
                    btn_vencimento.click()
            except:
                time.sleep(0.5)
            
        #Esperando o calendario aparecer
        element = driver.find_element(By.XPATH, campo_calendario)                                       
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
        caixa = driver.find_element(By.NAME, 'portador-bill')
        driver.execute_script("arguments[0].scrollIntoView();", caixa)
        caixa.send_keys(caixa_bling)
        caixa.send_keys(Keys.ENTER)
        #Baixando a conta
        categoria = driver.find_element(By.XPATH, "//*[contains(@id, 'categoria-duplicata-')]/div")
        categoria.click()
        texto_categoria = driver.find_element(By.ID, 'tree_seach')
        texto_categoria.send_keys(categoria_vendas)
        texto_categoria.send_keys(Keys.ENTER)
        btn_baixar_ok = driver.find_element(By.ID, 'btnBaixarBordero')
        btn_baixar_ok.click()
        #Retornando a tela de pedidos de venda
        driver.switch_to.window(driver.window_handles[0]) 
    df = pd.DataFrame(nao_encontrados)
    df.to_excel('Não encontrados.xlsx', index = False)


# In[64]:





