from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import funcoes
import re
import time


#importando a planilha
planilha = load_workbook("teste.xlsx")

#definindo a aba ativa
aba_ativa = planilha.active
#pagina = planilha['deals list']
pagina = planilha['companies']
    

# Percorrendo a planilha para adicionar na lista
empresasList = funcoes.buscar_cnpj(aba_ativa)

# Inicia o uma instância do google webdriver
service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Acessando a URL
url = 'https://cnpj.linkana.com/'

# Pesquisando os nomes das empresas da planilha e adicionando em uma lista
cnpj_list= []
for name in empresasList:
    driver.get(url)
    time.sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="q"]').click()
    driver.find_element(By.XPATH, '//*[@id="q"]').send_keys(name)
    driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/div[2]/form/div/input[2]').click()
    time.sleep(0.5)
    c = driver.find_element(By.XPATH, '/html/body/div/main/div/div/a/div/div/p[2]').text

    cnpj = re.sub(r'[^0-9]', '', c)

    cnpj_list.append(cnpj) #talvez substituir o nome da empresa na 'empresaList' pelo cnpj



empresas_result = funcoes.buscar_cnpj_api(cnpj_list)





print('ola mundo')