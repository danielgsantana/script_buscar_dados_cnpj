from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import funcoes
import re
import time


while True:
    try:
        print("Automação LDR")
        print("---------------------------------------------------")
        print("Escolha uma das opções abaixo:")
        print("1 - Para caso sua planilha conter os nomes da empresa")
        print("2 - Para caso sua planilha seja de CNPJ")
        entrada = int(input("-> "))

        if entrada == 1 or entrada == 2:
            break 
        else:
            print("---------------------------------------------------")
            print("Por favor, insira apenas o número 1 ou 2.")
            print("---------------------------------------------------")
            time.sleep(1)
    except ValueError:
        print("---------------------------------------------------")
        print("Entrada inválida. Por favor, insira apenas números.")
        print("---------------------------------------------------")
        time.sleep(1)


if entrada == 1:

    #importando a planilha (INSERIR O NOME DO ARQUIVO EXCEL QUE CONTENHA OS NOMES DAS EMPRESAS)
    planilha = load_workbook("nome empresas.xlsx")
    
    aba_ativa = planilha['principal']

    # Percorrendo a planilha para adicionar na lista
    empresasList = funcoes.capturar_cnpj(aba_ativa)

    cnpj_list = funcoes.buscar_cnpj(empresasList)

    empresas_result = funcoes.buscar_cnpj_api(cnpj_list)

    funcoes.adicionar_dados_planilha(empresasList, empresas_result)

if entrada == 2:

    #importando a planilha (INSERIR O NOME DO ARQUIVO EXCEL QUE CONTENHA OS NOMES DAS EMPRESAS)
    planilha = load_workbook("cnpj empresas.xlsx")
    
    aba_ativa = planilha['principal']

    # Percorrendo a planilha para adicionar na lista
    empresasList = funcoes.capturar_cnpj(aba_ativa)

    empresas_result = funcoes.buscar_cnpj_api(empresasList)

    funcoes.adicionar_dados_planilha(empresasList, empresas_result)
