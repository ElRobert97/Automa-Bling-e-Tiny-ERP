import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import locale
from time import sleep

# Define o locale para português
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')



def tela_notasfiscais(driver):
     while True:
        try:
            filtro = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "filter-clear")))
            driver.execute_script("arguments[0].click();", filtro)
            break
        except (TimeoutException, NoSuchElementException):
            sleep(2)
     #Trocando para pedidos de vendas
     driver.switch_to.window(driver.window_handles[1])
     

def tela_vendas(driver):
    #Tela de vendas
    while True:
        try:
            filtro = driver.find_element(By.CLASS_NAME, "filter-active" )
            filtro.click()
            data_ini = driver.find_element(By.ID, 'data-ini')
            por_periodo = driver.find_element(By.ID, 'opc-per-periodo').click()
            data_ini.clear()
            data_ini.send_keys('01/01/2024')
            btn_aplicar = driver.find_element(By.XPATH, '//*[@id="panel-vendas"]/div[1]/div[3]/ul/li[2]/div/div[5]/button[1] | //*[@id="panel-vendas"]/div[1]/div[3]/ul/li[1]/div/div[5]/button[1]').click()
            todas_vendas = driver.find_element(By.ID, 'sit-').click()
            break
        except:
            sleep(2)
    
def iniciar(login,senha, driver):
    driver.get('https://erp.tiny.com.br/vendas#list')
    driver.maximize_window()
    
    acesso_login = login
    acesso_senha = senha

    #Alerta do Tiny
    while True:
        try:
            alerta = driver.switch_to.alert
            if alerta:
                alerta.accept()   
                login = driver.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input').send_keys(acesso_login)
                senha = driver.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input').send_keys(acesso_senha)
                click = driver.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button').click()
                break
        except:
             login = driver.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input').send_keys(acesso_login)
             senha = driver.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input').send_keys(acesso_senha)
             click = driver.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button').click()
                
    

    tela_vendas(driver)
    #Abrir nova janela
    driver.execute_script("window.open('');")

    #Trocando para o contas a receber
    driver.switch_to.window(driver.window_handles[1])
    driver.get('https://erp.tiny.com.br/contas_receber')
    
    while True:
        try:
            sem_filtro = WebDriverWait(driver, 15).until(
                                    EC.visibility_of_element_located((By.CLASS_NAME, "filter-clear" )))
            driver.execute_script("arguments[0].click();", sem_filtro)
            filtro_historico = WebDriverWait(driver, 2).until(
                                    EC.visibility_of_element_located((By.CLASS_NAME, 'caret')))                      
            driver.execute_script("arguments[0].click();", filtro_historico)
            filtro_pesquisa = driver.find_elements(By.CLASS_NAME, 'no-icon')
            for i in filtro_pesquisa:
                if i.text == 'Histórico':
                    i.click() 
            driver.switch_to.window(driver.window_handles[0])    
            break
        except (TimeoutException, NoSuchElementException):
            sleep(2)
            
    # #Tela de notas fiscais
    # driver.execute_script("window.open('');")
    # driver.switch_to.window(driver.window_handles[2])
    # driver.get('https://erp.tiny.com.br/notas_fiscais#list')
    
    # while True:
    #     try:
    #         filtro = WebDriverWait(driver, 5).until(
    #             EC.element_to_be_clickable((By.CLASS_NAME, "filter-clear")))
    #         driver.execute_script("arguments[0].click();", filtro)
    #         break
    #     except (TimeoutException, NoSuchElementException):
    #         sleep(2)
    # #Trocando para pedidos de vendas
    # driver.switch_to.window(driver.window_handles[0])
           
def processar_pedidos(planilha, caixa_tiny, categoria , col_pedido, col_data, inicial, driver):
    nao_encontrados = {'Pedidos':[], 'Motivo':[]
                       }
    # Pré-localizar elementos fixos que não mudam com cada iteração da pagina contas a receber
    calendario_xpath = '/html/body/div[1]/div/div/div/div[2]/form/div[2]/div[2]/div/input'
    botao_receber_xpath = '/html/body/div[11]/ul/li[1]/a'
    salvar_bordero_xpath = '//*[@id="salvarBordero"]'
    
    
    
    for _, row in planilha.iterrows():
        pedido = inicial + row[col_pedido]
        data = row[col_data]
        
        #Aba pedidos de vendas
        barra_pesquisa = driver.find_element(By.XPATH, '//*[@id="pesquisa-mini"]')
        barra_pesquisa.clear()
        barra_pesquisa.send_keys(pedido)
        barra_pesquisa.send_keys(Keys.ENTER)
         
        try: 
            # Tentar encontrar o símbolo de contas a receber lançado
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//*[contains(@class, 'contas_lancadas_nota') or contains(@class, 'contas_lancadas')]")))
            driver.switch_to.window(driver.window_handles[1])
        except(NoSuchElementException, TimeoutException):
            #Realizar o lançamento da conta a receber
            try:
                encontrado = False
                #Clickar nos 3 pontos para abrir as opções
                tres_pontos = WebDriverWait(driver, 2).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'button-navigate')))
                tres_pontos.click()
                
                #Acessando o menu
                menu = driver.find_elements(By.TAG_NAME, 'a')

                #Iterando sobre as opções
                for opcao in menu:
                    if 'lançar contas' in opcao.text:
                        opcao.click()
                        encontrado = True
                        try:
                            fechar = WebDriverWait(driver, 2).until(
                            EC.visibility_of_element_located((By.CLASS_NAME, 'hotkey-label')))
                            fechar.click()
                        except:
                            pass
                        driver.switch_to.window(driver.window_handles[1])               
                        break
                #Se não encontrar essas informações no loop vai procurar em notas fiscais
                if not encontrado:
                    element = driver.find_element(By.CLASS_NAME, "nota_gerada")
                    element.click()
                    tres_pontos = WebDriverWait(driver, 10).until(
                                    EC.presence_of_all_elements_located((By.XPATH, '//button')))
                    for i in tres_pontos:
                        if 'mais ações' in i.text or 'ações' in i.text:
                            i.click()
                            break
                    menu =  WebDriverWait(driver, 5).until(
                                    EC.presence_of_all_elements_located((By.TAG_NAME, 'a')))
                    for opcao in menu:
                        if 'lançar contas' in opcao.text:
                            opcao.click()
                            break       
                driver.get('https://erp.tiny.com.br/vendas#list')
                while True:
                    try:
                        filtro = driver.find_element(By.CLASS_NAME, "filter-active" )
                        filtro.click()
                        data_ini = driver.find_element(By.ID, 'data-ini')
                        por_periodo = driver.find_element(By.ID, 'opc-per-periodo').click()
                        data_ini.clear()
                        data_ini.send_keys('01/01/2024')
                        btn_aplicar = driver.find_element(By.XPATH, '//*[@id="panel-vendas"]/div[1]/div[3]/ul/li[2]/div/div[5]/button[1]').click()
                        break
                    except:
                        sleep(2)
                driver.switch_to.window(driver.window_handles[1])         
            except (TimeoutException, NoSuchElementException):
                nao_encontrados['Pedidos'].append(pedido)
                texto = "Não foi possivel lançar a conta"
                nao_encontrados['Motivo'].append(texto)
                pass
                    
        
        # Limpar e inserir o pedido na barra de pesquisa
        barra_pesquisa = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="pesquisa-mini"]')))
        barra_pesquisa.clear()
        barra_pesquisa.send_keys(pedido)
        barra_pesquisa.send_keys(Keys.ENTER)
        
        try:          
                element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, '//*[contains(@id, "linha")]/td[2]/button'))
                                                    )
                driver.execute_script("arguments[0].scrollIntoView();", element)
                driver.execute_script("arguments[0].click();", element)
                
                # Verifica se o botão de receber está presente e clica nele
                botao_receber = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, botao_receber_xpath))
                )
                botao_receber.click()
                    
                caixa = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located(((By.ID, 'idContaContabil')))
                )
                
                # Selecionar o caixa
                caixa = Select(driver.find_element(By.ID, 'idContaContabil'))
                caixa.select_by_visible_text(caixa_tiny)
                
                # Selecionar "Vendas E-commerce" na categoria
                categoria_vendas = Select(driver.find_element(By.ID, 'idCategoria'))
                categoria_vendas.select_by_visible_text(categoria)
                
                # Limpar e inserir a data no calendário
                calendario = driver.find_element(By.XPATH, calendario_xpath)
                calendario.clear()
                calendario.send_keys(data)
                
                # Clicar em "Salvar"
                driver.find_element(By.XPATH, salvar_bordero_xpath).click()
                driver.switch_to.window(driver.window_handles[0])           
        except (NoSuchElementException, TimeoutException):
            # Verificar se a página retornou resultados 
            driver.switch_to.window(driver.window_handles[0])
            continue
    df_naoencontrados = pd.DataFrame(nao_encontrados)
    df_naoencontrados.to_excel('Pedidos nao encontrados.xlsx', index=False)
    
