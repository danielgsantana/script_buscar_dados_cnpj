from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import requests
import time
import json
import re

import Empresa

#percorrendo a planilha para adicionar na lista
def capturar_cnpj (list) :
    
    empresasList = []
    for celula in list["A"]:
        empresasList.append(celula.value)

    #del empresasList[0]
    #del empresasList[0]

    return empresasList

def buscar_cnpj (empresasList) :

    # Inicia o uma instância do google webdriver
    service = Service()
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)

    # Acessando a URL
    url = 'https://cnpj.linkana.com/'

    # Pesquisando os nomes das empresas da planilha e adicionando em uma lista
    cnpj_list= []

    driver.get(url)
    time.sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="q"]').click()
    driver.find_element(By.XPATH, '//*[@id="q"]').send_keys('azship')
    driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/div[2]/form/div/input[2]').click()
    time.sleep(0.5)

    for name in empresasList:
        
        driver.find_element(By.XPATH, '//*[@id="q"]').click()
        driver.find_element(By.XPATH, '//*[@id="q"]').send_keys(name)
        driver.find_element(By.XPATH, '/html/body/div/main/div/form/div[2]/input').click()
        time.sleep(0.5)
        
        try:
            c = driver.find_element(By.XPATH, '/html/body/div/main/div/div/a/div/div/p[2]').text
            cnpj = re.sub(r'[^0-9]', '', c)
            cnpj_list.append(cnpj)

        except NoSuchElementException:
            print(f"A empresa {name} não pode ser encontrada pelo nome!")
            c = "CNPJ não encontrado"     # parei aqui, não posso adicionar uma string em uma lista de inteiro
            cnpj_list.append(c)

    return cnpj_list

# Loop sobre os CNPJs
def buscar_cnpj_api(cnpj_list) :
    
    # Substitua "seu_token_aqui" pelo seu token real
    token = "f2987ff033bd4550a7a208f4fc82be13c3ea145f2369561b089cba941742eda6"
    
    # Configurar os headers com o token
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    dados_cnpj = []

    response_cnpj = []
    for cnpj in cnpj_list:
        

        # Verificar se a requisição foi bem-sucedida (código 200)
        if response.status_code == 200:

            if isinstance(cnpj, str):
                    
                objeto = Empresa.Empresa(cnpj= "0",
                                razao_social= "0",
                                nome_fantasia= "0",
                                abertura= "0",
                                capital= "0",
                                email= "0",
                                telefone= "0",
                                municipio= "0",
                                uf= "0",
                                cep= "0",
                                cnae =  "0")

                response_cnpj.append(objeto)
            else:
                # Processar os dados da resposta
                response = response.json()

                objeto = Empresa.Empresa(cnpj=response['cnpj'],
                                razao_social=response['nome'],
                                nome_fantasia=response['fantasia'],
                                abertura=response['abertura'],
                                capital=response['capital_social'],
                                email=response['email'],
                                telefone=response['telefone'],
                                municipio=response['municipio'],
                                uf=response['uf'],
                                cep=response['cep'],
                                cnae = response.get('atividade_principal', [{}])[0].get('text', ''))

                response_cnpj.append(objeto)

        if cnpj == "CNPJ não encontrado":
            c = Empresa.Empresa(cnpj='erro',
                                razao_social= 'erro',
                                nome_fantasia= 'erro',
                                abertura= 'erro',
                                capital= 'erro',
                                email= 'erro',
                                telefone= 'erro',
                                municipio= 'erro',
                                uf= 'erro',
                                cep= 'erro',
                                cnae = 'erro')
            
            response_cnpj.append(c)
        else:
            # Construa a URL com a variável 'cnpj'
            url = f"https://receitaws.com.br/v1/cnpj/{cnpj}/days/5"

            # Fazer a requisição GET
            response = requests.get(url, headers=headers)


            # Verificar se a requisição foi bem-sucedida (código 200)
            if response.status_code == 200:
                # Processar os dados da resposta
                response = response.json()

                objeto = Empresa.Empresa(cnpj=response['cnpj'],
                                razao_social=response['nome'],
                                nome_fantasia=response['fantasia'],
                                abertura=response['abertura'],
                                capital=response['capital_social'],
                                email=response['email'],
                                telefone=response['telefone'],
                                municipio=response['municipio'],
                                uf=response['uf'],
                                cep=response['cep'],
                                cnae = response.get('atividade_principal', [{}])[0].get('text', ''))

                response_cnpj.append(objeto)

            else:
                # Lidar com erros
                print(f"Erro para CNPJ {cnpj}: {response.status_code} - {response.text}")

    return response_cnpj

def adicionar_dados_planilha (empresasList, empresas_result) :
    
    planilha = Workbook()
    aba_ativa = planilha.active

    aba_ativa['A1'] = 'NOME EMPRESA'
    aba_ativa['B1'] = 'CNPJ'
    aba_ativa['C1'] = 'RAZÃO SOCIAL'
    aba_ativa['D1'] = 'NOME FANTASIA'
    aba_ativa['E1'] = 'ABERTURA'
    aba_ativa['F1'] = 'CAPITAL'
    aba_ativa['G1'] = 'EMAIL'
    aba_ativa['H1'] = 'TELEFONE'
    aba_ativa['I1'] = 'MUNICIPIO'
    aba_ativa['J1'] = 'UF'
    aba_ativa['K1'] = 'CEP'
    aba_ativa['L1'] = 'CNAE PRINCIPAL'

    row = 2
    repetidor = 0

    for rows in empresas_result :
        aba_ativa[f'A{row}'] = empresasList[repetidor]
        aba_ativa[f'B{row}'] = empresas_result[repetidor].cnpj
        aba_ativa[f'C{row}'] = empresas_result[repetidor].razao_social
        aba_ativa[f'D{row}'] = empresas_result[repetidor].nome_fantasia
        aba_ativa[f'E{row}'] = empresas_result[repetidor].abertura
        aba_ativa[f'F{row}'] = empresas_result[repetidor].capital
        aba_ativa[f'G{row}'] = empresas_result[repetidor].email
        aba_ativa[f'H{row}'] = empresas_result[repetidor].telefone
        aba_ativa[f'I{row}'] = empresas_result[repetidor].municipio
        aba_ativa[f'J{row}'] = empresas_result[repetidor].uf
        aba_ativa[f'K{row}'] = empresas_result[repetidor].cep
        aba_ativa[f'L{row}'] = empresas_result[repetidor].cnae

        repetidor += 1 
        row += 1

    planilha.save('Dados_Gerados.xlsx')
