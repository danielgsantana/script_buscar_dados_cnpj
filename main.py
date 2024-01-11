from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import funcoes


# Substitua "seu_token_aqui" pelo seu token real
token = "f2987ff033bd4550a7a208f4fc82be13c3ea145f2369561b089cba941742eda6"


#importando a planilha
planilha = load_workbook("teste.xlsx")

#definindo a aba ativa
aba_ativa = planilha.active
#pagina = planilha['deals list']
#pagina = planilha['companies']
    

#percorrendo a planilha para adicionar na lista
empresasList = funcoes.buscar_cnpj(aba_ativa)

#Inicia o uma instância do google webdriver
service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

#acessando a URL
url = 'https://cnpj.linkana.com/'
driver.get(url)


fake_cnpj = []

#Realizando a pesquisa
driver.find_element(By.XPATH, '//*[@id="q"]').click()
driver.find_element(By.XPATH, '//*[@id="q"]').send_keys(empresasList[0])
driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/div[2]/form/div/input[2]').click()


#fake_cnpj [0] = empresasList[0]
qtdEmpresas = len(empresasList)

empresas_result = []

for i in range(qtdEmpresas):
    empresas_result.append(funcoes.buscar_cnpj_api(qtdEmpresas, empresasList, token))

#print(response)

empresa1 = empresas_result[0] # tratar os dados para somente oq é necessario adicionar na planilha

print(empresa1['nome'])
print(empresas_result[1])
print(empresas_result[2].nome)