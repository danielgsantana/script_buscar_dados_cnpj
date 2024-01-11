from openpyxl import Workbook, load_workbook
import requests
import json


#percorrendo a planilha para adicionar na lista
def buscar_cnpj (list):
    
    empresasList = []
    for celula in list["A"]:
        empresasList.append(celula.value)

    #del empresasList[0]
    #del empresasList[0]

    return empresasList


# Loop sobre os CNPJs
def buscar_cnpj_api(quantidade_cnpj, cnpj_list, token):
    
    dados_cnpj = []
    
    
    # Configurar os headers com o token
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    contador = 0

    for cnpj in range(quantidade_cnpj):

        # Construa a URL com a variável 'cnpj'
        url = f"https://receitaws.com.br/v1/cnpj/{cnpj_list[contador]}/days/5"

        # Fazer a requisição GET
        response = requests.get(url, headers=headers)

        # Verificar se a requisição foi bem-sucedida (código 200)
        if response.status_code == 200:
            # Processar os dados da resposta
            response = response.json()
            print(f"CNPJ:{response}")

            contador = contador + 1
            #dados_cnpj.append(response)

            return response
        else:
            contador = contador + 1

            # Lidar com erros
            print(f"Erro para CNPJ {cnpj}: {response.status_code} - {response.text}")


def tratar_dados_api (objetos_api):
    objeto_tratado = {
        "cnpj": None,
        "Razão Social": None,
        "Nome Fantasia": None,
        "Situação Cadastral": None,
        "Data de Abertura": None,
        "Capital Social": None,
        "Email": None,
        "Telefone": None,
        "Município": None,
        "UF": None,
        "CEP": None,
        "CNAE Principal": None,
    }

    objeto_tratado.cnpj =  objetos_api.