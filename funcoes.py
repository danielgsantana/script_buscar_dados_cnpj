from openpyxl import Workbook, load_workbook
import requests
import json
import Empresa

#percorrendo a planilha para adicionar na lista
def buscar_cnpj (list):
    
    empresasList = []
    for celula in list["A"]:
        empresasList.append(celula.value)

    #del empresasList[0]
    #del empresasList[0]

    return empresasList


# Loop sobre os CNPJs
def buscar_cnpj_api(cnpj_list):
    
    # Substitua "seu_token_aqui" pelo seu token real
    token = "f2987ff033bd4550a7a208f4fc82be13c3ea145f2369561b089cba941742eda6"
    
    # Configurar os headers com o token
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    dados_cnpj = []

    for cnpj in cnpj_list:

        # Construa a URL com a variável 'cnpj'
        url = f"https://receitaws.com.br/v1/cnpj/{cnpj}/days/5"

        # Fazer a requisição GET
        response = requests.get(url, headers=headers)

        response_cnpj = []
        # Verificar se a requisição foi bem-sucedida (código 200)
        if response.status_code == 200:
            # Processar os dados da resposta
            response = response.json()
            
            response_cnpj.append(Empresa.Empresa(cnpj=response['cnpj'],
                            razao_social=response['nome'],
                            nome_fantasia=response['fantasia'],
                            abertura=response['abertura'],
                            capital=response['capital_social'],
                            email=response['email'],
                            telefone=response['telefone'],
                            municipio=response['municipio'],
                            uf=response['uf'],
                            cep=response['cep'],
                            cnae = response.get('atividade_principal', [{}])[0].get('text', '')))
            
            
            
        else:
            # Lidar com erros
            print(f"Erro para CNPJ {cnpj}: {response.status_code} - {response.text}")

    return response_cnpj

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
