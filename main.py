from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import funcoes
import re
import time


#importando a planilha (INSERIR O NOME DO ARQUIVO EXCEL QUE CONTENHA OS NOMES DAS EMPRESAS)
planilha = load_workbook("teste.xlsx")
 
aba_ativa = planilha.active

# Percorrendo a planilha para adicionar na lista
empresasList = funcoes.capturar_cnpj(aba_ativa)


cnpj_list = funcoes.buscar_cnpj(empresasList)


empresas_result = funcoes.buscar_cnpj_api(cnpj_list)


funcoes.adicionar_dados_planilha(empresasList, empresas_result)


print('ola mundo')