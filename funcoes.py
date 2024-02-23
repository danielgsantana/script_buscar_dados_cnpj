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
        
        if celula.value == None:
            break
            
        else:
            caracteres_a_manter = r"a-zA-ZáéíóúÁÉÍÓÚãõÃÕçÇ.\/1234567890" 

            expressao_regular = f"[^{caracteres_a_manter} ]"
            
            # Removendo caracteres especiais usando expressões regulares
            d = re.sub(expressao_regular, '', celula.value)
            
            # Substituindo dois espaços consecutivos por apenas um espaço
            empresa = re.sub(r'  +', ' ', d)



        empresasList.append(empresa)


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

def fazer_requisicao (cnpj) :
    
    # Substitua "seu_token_aqui" pelo seu token real
    token = "f2987ff033bd4550a7a208f4fc82be13c3ea145f2369561b089cba941742eda6"
    
    # Configurar os headers com o token
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # Construa a URL com a variável 'cnpj'
    url = f"https://receitaws.com.br/v1/cnpj/{cnpj}/days/5"
       
    try:
        # Fazer a requisição GET
        response = requests.get(url, headers=headers)
        time.sleep(2)
        response.raise_for_status()
        return response.json()  # Retorna os dados da resposta em formato JSON
    
    except requests.exceptions.RequestException as e:
        print(f"Erro ao fazer requisição: {e}")
        return None

# Loop sobre os CNPJs
def buscar_cnpj_api(cnpj) :

    if cnpj == "CNPJ não encontrado":
        c = Empresa.Empresa(cnpj='Empresa não encontrada pelo nome',
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
        
        return c
    else:

        while True: 
            response_json = fazer_requisicao(cnpj)

            if response_json:
                # Se a resposta foi recebida com sucesso, você pode continuar com o processamento
                
                c = Empresa.Empresa(cnpj=response_json['cnpj'],
                           razao_social=response_json['nome'],
                           nome_fantasia=response_json['fantasia'],
                           abertura=response_json['abertura'],
                           capital=response_json['capital_social'],
                           email=response_json['email'],
                           telefone=response_json['telefone'],
                           municipio=response_json['municipio'],
                           uf=response_json['uf'],
                           cep=response_json['cep'],
                           cnae = response_json.get('atividade_principal', [{}])[0].get('text', ''))
           
                return c
            
            else:
                # Se houve um erro na requisição, espere um pouco antes de tentar novamente
                print("Aguardando resposta da API...")
                time.sleep(5)  # Espera 5 segundos antes de tentar novamente 

def adicionar_dados_planilha (nome_empresa, empresas_result, planilha, row) :
    
    
    aba_ativa = planilha['resultado gerado']
    
    aba_ativa[f'A{row}'] = nome_empresa
    aba_ativa[f'B{row}'] = empresas_result.cnpj
    aba_ativa[f'C{row}'] = empresas_result.razao_social
    aba_ativa[f'D{row}'] = empresas_result.nome_fantasia
    aba_ativa[f'E{row}'] = empresas_result.abertura
    aba_ativa[f'F{row}'] = empresas_result.capital
    aba_ativa[f'G{row}'] = empresas_result.email
    aba_ativa[f'H{row}'] = empresas_result.telefone
    aba_ativa[f'I{row}'] = empresas_result.municipio
    aba_ativa[f'J{row}'] = empresas_result.uf
    aba_ativa[f'K{row}'] = empresas_result.cep
    aba_ativa[f'L{row}'] = empresas_result.cnae

    planilha.save('planilha.xlsx')