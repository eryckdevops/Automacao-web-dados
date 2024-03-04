from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import pyautogui
import time

driver_path = ''
file_path = ''
url_form = ""

s = Service(driver_path)
driver = webdriver.Chrome(service=s)
driver.maximize_window()

df = pd.read_excel(file_path)

driver.get(url_form)
time.sleep(8) 

for index, row in df.iterrows():
    try:
        time.sleep(8)  
        campo_assunto = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'Subject')))
        campo_assunto.clear()
        campo_assunto.send_keys(row['Status'])
        
        time.sleep(8)  
        
        campo_nome = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'Name')))
        campo_nome.clear()
        campo_nome.send_keys(row['Organização'])
        
        time.sleep(8) 
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-container"))).click()
        prioridade_valor = row['Prioridade']
        prioridade_texto = {
            "Alta": "Alta",
            "Baixa": "Baixa",
            "Média": "Média",
            "Urgente": "Urgente"
        }.get(prioridade_valor, "- Selecione -")
        opcao_urgencia_xpath = f"//div[@class='select2-result-label' and text()='{prioridade_texto}']"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, opcao_urgencia_xpath))).click()
        
        campo_servico = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".md-select-treeview-dropdown-field")))
        campo_servico.click()
        time.sleep(8) 
        
        tipo_valor = row['Tipo']
        tipo_to_xpath = {
            "Desenvolvimento de Personalização": "//div[contains(text(), 'Desenvolvimento de Personalização')]",
            "Produto": "//div[contains(text(), 'Produto')]",
            "Serviço de Implantação": "//div[contains(text(), 'Serviço de Implantação')]",
            "Suporte": "//div[contains(text(), 'Suporte')]",
        }
        
        opcao_servico_xpath = tipo_to_xpath.get(tipo_valor, None)
        
        if opcao_servico_xpath:
            opcao_servico = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, opcao_servico_xpath)))
            opcao_servico.click()
        else:
            print(f"Tipo '{tipo_valor}' não encontrado nas opções de serviço.")
        
        time.sleep(8) 
        
        campo_email = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'Email')))
        campo_email.clear()
        campo_email.send_keys(row['Email'])
        
        time.sleep(8) 
        
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "Description_ifr")))
        descricao = f"Status: {row['Status']}, Agente: {row['Agente']}, Criação: {row['Criação']}, Respondido: {row['Respondido']}, Categoria: {row['Categoria']}"
        campo_descricao = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "tinymce")))
        driver.execute_script("arguments[0].innerText = arguments[1]", campo_descricao, descricao)
        driver.switch_to.default_content()
        
        time.sleep(5)
        botao_posicao_x = 522
        botao_posicao_y = 478
        time.sleep(2) 
        pyautogui.moveTo(botao_posicao_x, botao_posicao_y) 
        pyautogui.doubleClick(button='left') 
        
        time.sleep(8)  
        
        botao_posicao_x = 950
        botao_posicao_y = 644
        time.sleep(4) 
        pyautogui.moveTo(botao_posicao_x, botao_posicao_y)  
        pyautogui.click(button='left') 
        
        time.sleep(8) 
        


        time.sleep(4)
        
        driver.refresh()
        time.sleep(8)  
    except TimeoutException as e:
        print(f"Erro de Timeout: {e}")
    except Exception as e:
        print(f"Ocorreu um erro ao preencher ou enviar o formulário: {e}")
    finally:
        pass

driver.quit()
