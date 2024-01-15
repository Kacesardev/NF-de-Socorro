from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options



import traceback
import time
import os
import shutil
import sys



destino = r"\\172.10.10.5\itaipu-fs\Interdep\CSC-Contabilidade\FISCAL ITAIPU\Socorro\Base_Dados_Relatorio_Sicon_teste\\"

nome_padrao_arquivo = "NFEmitida.xml"
novo_nome_arquivo = "NotaEmitida.xml"
time_out_download_arquivo = 120 #segundos

nome_padrao_arquivo2 = "RelatorioRelacaoOrdensServico.slk"
novo_nome_arquivo2 = "RelatorioOS.slk"
time_out_download_arquivo = 120 #segundos

nome_padrao_arquivo3 = "Garantia.xml"
novo_nome_arquivo3 = "GarantiaEmitida.xml"
time_out_download_arquivo = 120 #segundos

options = Options()

options.add_experimental_option("prefs", {
  "download.default_directory": destino,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "plugins.plugins_disabled": ["Chrome PDF Viewer"],
  "safebrowsing.enabled": True,
  "safebrowsing.disable_download_protection": True,
  "safebrowsing-disable-auto-update": True,
  "disable-client-side-phishing-detection": True
})



try:
    # Configuração do serviço e abertura do navegador
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico,options=options)

    navegador.get("https://acp-siconnet.scania.com.br/sicomnet-ace3/wlm/sicomweb.gen.gen.pag.Login.cls")

    # Espera até que o alerta esteja presente
    WebDriverWait(navegador, 10).until(EC.alert_is_present())

    # Aceita o alerta
    alerta = Alert(navegador)
    alerta.accept()

    # Efetua login
    login = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, 'control_11')))
    login.send_keys('WLMCPJ')

    senha = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, 'control_15')))
    senha.send_keys('Dez,0119')

    botao_login = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.ID, 'botaoLogin')))
    botao_login.click()

    time.sleep(5)

    # Entra na rotina de Notas Fiscais (Rotina 1)

    def entra_notas_emitidas(driver):
        print("Entrando em Notas Emitidas.\n")
        aba_original = driver.window_handles[1]
        driver.switch_to.window(aba_original)
        javascript_code = "zenPage.carregaFrame('sicomweb.ccr.fr.pgr.NFEmitida.cls', 125, 0);"
        driver.execute_script(javascript_code)

        return driver
    
    # Seleciona o dropdown, clica e preenche os campos como: Filial, Operação, Data Inicial/Final e os demais campos marcados.
    def seleciona_dropdown(driver):
        print("Preenchendo campo de filial com o número 1.\n")
        time.sleep(5)

    # Espera até que o iframe esteja presente e muda para o iframe
        iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'iframe_34')))
        driver.switch_to.frame(iframe)

        
        driver.find_element(By.ID, 'input_13').click()
        driver.find_element(By.ID, 'input_13').send_keys('1')
        driver.find_element(By.ID, 'control_18').click()
        driver.find_element(By.ID, 'control_18').send_keys('CP,GS,PP,VG,VM')
        driver.find_element(By.ID, 'control1_19').click()
        driver.find_element(By.ID, 'control1_19').send_keys('H-31')
        driver.find_element(By.ID, 'control2_19').click()
        driver.find_element(By.ID, 'control2_19').send_keys('H-1')
        driver.find_element(By.ID, 'textRadio_2_25').click()
        driver.find_element(By.ID, 'textRadio_2_31').click()
        driver.find_element(By.ID, 'textRadio_1_32').click()
        driver.find_element(By.ID, 'textRadio_2_26').click()
        
        
        javascript_code2 = "zenPage.getComponent(35).showDropdown();"
        driver.execute_script(javascript_code2)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_35').is_selected()
        
        driver.find_element(By.ID,'control_36').click()
        driver.find_element(By.ID,'control_36').is_selected()
        driver.find_element(By.ID, 'control_39').click()

        

        # Aguardar janela de donwload
        while (len(driver.window_handles)<3):
            pass

        driver.switch_to.window(driver.window_handles[2]) #janela de download

        time.sleep(1)   
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//tr//span[@class='iconDownloadBase']")))
        #clica 2 vezes pra fazer o download
        element.click()
        time.sleep(1)
        element.click()   

        # Aguardar donwload do arquivo....
        waits = 0
        while not os.path.isfile(destino + nome_padrao_arquivo) and waits < time_out_download_arquivo:
            time.sleep(1)
            waits +=1
        # Renomear arquivo
        try:
            shutil.move(destino + nome_padrao_arquivo, destino + novo_nome_arquivo)
        except:
            print("Não foi possível renomear o arquivo " + nome_padrao_arquivo + "Mensagem: " + sys.exc_info()[0])
        waits = 0
        while not os.path.isfile(destino + novo_nome_arquivo) and waits < time_out_download_arquivo:
            time.sleep(3)
            waits +=1
        driver.switch_to.window(driver.window_handles[3])
        driver.close()
        driver.switch_to.window(driver.window_handles[2])
        driver.close() 
        driver.switch_to.window(driver.window_handles[1])

        # Espera até que o iframe esteja presente e muda para o iframe
        iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'iframe_34'))).click()
        driver.switch_to.frame(iframe)

        time.sleep(2)
       
       #Entrando em Relatorio de OS (Rotina 2)
        
        driver.find_element(By.ID, 'a9').click() # Comercial
        driver.find_element(By.ID, 'm60').click() # Oficina
        driver.find_element(By.ID, 'g185').click() # Relatório
        driver.find_element(By.XPATH, '//*[@id="rotina"]/table[1]/tbody/tr[19]/td/a').click() # Relação de Ordens de Serviço
        time.sleep(1)

        iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'iframe_34')))
        driver.switch_to.frame(iframe)
        

        
        # Clicar e preencher o campo filial
        driver.find_element(By.XPATH, '//*[@id="input_14"]').send_keys('1')

        # Seleciona Tipo de  Relatorio 
        javascript_code3 = "zenPage.getComponent(15).showDropdown();"
        driver.execute_script(javascript_code3)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_15').is_selected()

        time.sleep(1)
        
        # Seleciona Tipo de Layout
        javascript_code4 = "zenPage.getComponent(17).showDropdown();"
        driver.execute_script(javascript_code4)
        ActionChains(driver).send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_17').is_selected()

       
        
       # Opção de Seleção
        javascript_code5 = "zenPage.getComponent(19).showDropdown();"
        driver.execute_script(javascript_code5)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_19').is_selected()

        
        
       # Seleção Adicional
        
        javascript_code6 = "zenPage.getComponent(21).showDropdown();"
        driver.execute_script(javascript_code6)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_21').is_selected()

        
        
        # Data Inicial
        driver.find_element(By.ID,'control1_23').click()
        driver.find_element(By.ID,'control1_23').send_keys('H-31')
        # Data Final
        driver.find_element(By.ID, 'control2_23').click()
        driver.find_element(By.ID, 'control2_23').send_keys('H-1')

        #Classificação por Tipo
        driver.find_element(By.XPATH, '//*[@id="textRadio_2_28"]').click()

                
        # Seleciona classificação principal/secundaria
        javascript_code7 = "zenPage.getComponent(29).showDropdown();"
        driver.execute_script(javascript_code7)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_29').is_selected()

        

        # Opção de Seleção Adicional
        javascript_code8 = "zenPage.getComponent(30).showDropdown();"
        driver.execute_script(javascript_code8)
        ActionChains(driver).send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_30').is_selected()
        

        # Valores
        javascript_code9 = "zenPage.getComponent(39).showDropdown();"
        driver.execute_script(javascript_code9)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_39').is_selected()
        

        # Formato do Relatório
        javascript_code10 = "zenPage.getComponent(52).showDropdown();"
        driver.execute_script(javascript_code10)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_52').is_selected()


        driver.find_element(By.ID, 'control_56').click() #Processar

        # Aguardar janela de donwload
        while (len(driver.window_handles)<3):
            pass

        
            
     # Aguardar donwload do arquivo....
        waits = 0
        while not os.path.isfile(destino + nome_padrao_arquivo2) and waits < time_out_download_arquivo:
            time.sleep(2)
            waits +=1
        # Renomear arquivo
        try:
            shutil.move(destino + nome_padrao_arquivo2, destino + novo_nome_arquivo2)
        except:
            print("Não foi possível renomear o arquivo " + nome_padrao_arquivo + "Mensagem: " + sys.exc_info()[0])
        waits = 0
        while not os.path.isfile(destino + novo_nome_arquivo2) and waits < time_out_download_arquivo:
            time.sleep(3)
            waits +=1
        driver.switch_to.window(driver.window_handles[2])
        driver.close()
        driver.switch_to.window(driver.window_handles[1])

        
        

        # Espera até que o iframe esteja presente e muda para o iframe
        iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'iframe_34'))).click()
        driver.switch_to.frame(iframe)

        time.sleep(1)

        
      
                                      
        #Tira Relatorio de Garantia( Rotina 3 )

        driver.find_element(By.ID, 'a9').click() # Comercial
        driver.find_element(By.ID, 'm63').click() # Garantia
        driver.find_element(By.ID, 'g196').click() # Relatório/Arquivo
        driver.find_element(By.XPATH, '//*[@id="rotina"]/table/tbody/tr[8]/td/a').click() #Garantia
        time.sleep(1)                                            
    
    # Entrar no iframe
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'iframe_34')))

        time.sleep(2)

        # Clicar e preencher o campo filial
        driver.find_element(By.XPATH, '//*[@id="input_14"]').send_keys('1')

        # Opção de Seleção
        javascript_code11 = "zenPage.getComponent(15).showDropdown();"
        driver.execute_script(javascript_code11)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_15').is_selected()
        

        # Data Inicial
        driver.find_element(By.ID,'control1_16').click
        driver.find_element(By.ID,'control1_16').send_keys('H-31')
        # Data Final
        driver.find_element(By.ID, 'control2_16').click()
        driver.find_element(By.ID, 'control2_16').send_keys('H-1')

        

        #Opção de Classificação
        javascript_code12 = "zenPage.getComponent(24).showDropdown();"
        driver.execute_script(javascript_code12)
        ActionChains(driver).send_keys("\ue015")
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_24').is_selected()
        

        # Lista Itens
        javascript_code13 = "zenPage.getComponent(26).showDropdown();"
        driver.execute_script(javascript_code13)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_26').is_selected()
       


        #Formato do Relatório
        javascript_code14 = "zenPage.getComponent(33).showDropdown();"
        driver.execute_script(javascript_code14)
        ActionChains(driver).send_keys("\ue015").send_keys("\ue015").perform()
        ActionChains(driver).send_keys("\ue007").perform()
        driver.find_element(By.ID, 'btn_33').is_selected()
       
        
        
        driver.find_element(By.ID, 'control_34').click()
        time.sleep(1)
        driver.find_element(By.ID, 'control_37').click()


# Aguardar janela de donwload
        while (len(driver.window_handles)<3):
            pass

        driver.switch_to.window(driver.window_handles[2]) #janela de download

        time.sleep(1)   
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//tr//span[@class='iconDownloadBase']")))
        #clica 2 vezes pra fazer o download
        element.click()
        time.sleep(1)
        element.click()   



# Aguardar donwload do arquivo....
        waits = 0
        while not os.path.isfile(destino + nome_padrao_arquivo3) and waits < time_out_download_arquivo:
            time.sleep(2)
            waits +=1
        # Renomear arquivo
        try:
            shutil.move(destino + nome_padrao_arquivo3, destino + novo_nome_arquivo3)
        except:
            print("Não foi possível renomear o arquivo " + nome_padrao_arquivo + "Mensagem: " + sys.exc_info()[0])
        waits = 0
        while not os.path.isfile(destino + novo_nome_arquivo2) and waits < time_out_download_arquivo:
            time.sleep(3)
            waits +=1
        driver.switch_to.window(driver.window_handles[3])
        driver.close()
        driver.switch_to.window(driver.window_handles[2])
        driver.close()
        driver.switch_to.window(driver.window_handles[1])
        driver.quit()
        
    # Chama a função
    navegador = entra_notas_emitidas(navegador)
    navegador = seleciona_dropdown(navegador)
           
    
except Exception as e:
    # Imprime a exceção em caso de erro
    print(f"Erro: {str(e)}")
    traceback.print_exc()  # Imprime o traceback para ajudar a identificar a causa do erro

