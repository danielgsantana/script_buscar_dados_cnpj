from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

#importando a planilha
planilha = load_workbook("basePY.xlsx")

#definindo a aba ativa
aba_ativa = planilha.active
pagina = planilha['deals list']

#percorrendo a planilha para adicionar na lista (pegar da posição 2 em diante)
empresasList = []
for celula in aba_ativa["A"]:
    empresasList.append(celula.value)
del empresasList[0]
del empresasList[0]

#print(empresasList)

#Inicia o uma instância do google webdriver
service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

#acessando a URL
url = 'https://cnpj.linkana.com/'
driver.get(url)

#Realizando a pesquisa
driver.find_element(By.XPATH, '//*[@id="q"]').click()
driver.find_element(By.XPATH, '//*[@id="q"]').send_keys(empresasList[0])
driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/div[2]/form/div/input[2]').click()

filiais = driver.find_elements(By.TAG_NAME, 'a')
#del filiais[0]
#print(len(filiais))


dados_filiais = []
for i in range(filiais):
    new_var = range
    driver.find_elements(By.TAG_NAME, 'a')[filiais[new_var]].click()

    #Registrando em variaveis o caminho para capturarmos os dados
    cnpj = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[1]/div/h2[2]/b[2]').text
    razao_social = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[1]/li[1]/p').text
    nome_fantasia = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[1]/li[2]/p').text
    situacao_cadastral = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[1]/li[3]/p').text
    data_abertura = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[1]/li[5]/p').text
    capital_social = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[1]/li[7]/p').text
    email = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[1]/li[11]/p').text
    telefone = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[1]/li[12]/p').text
    municipio = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[2]/li[5]/p').text
    uf = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[2]/li[6]/p').text
    cep = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/ul[2]/li[7]/p').text
    cnae = driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[8]/ul/li[2]').text

    dados_filiais.append(cnpj, razao_social, nome_fantasia, situacao_cadastral, 
                         data_abertura, capital_social, email, telefone, municipio, uf, cep, cnae) 

planilha.append(dados_filiais)


#Clicando no primeiro card
#driver.find_element(By.XPATH, '/html/body/div[1]/main/div/div/a/div/div/p[1]').click()











print('ola mundo')

