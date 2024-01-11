import requests

# Loop sobre os CNPJs
def buscar_cnpj_api(cnpj_list, token):
    # Configurar os headers com o token
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    for cnpj in cnpj_list:
        # Construa a URL com a variável 'cnpj'
        url = f"https://receitaws.com.br/v1/cnpj/{cnpj}/days/5"

        # Fazer a requisição GET
        response = requests.get(url, headers=headers)

        # Verificar se a requisição foi bem-sucedida (código 200)
        if response.status_code == 200:
            # Processar os dados da resposta
            data = response.json()
            print(f"Resultados para CNPJ {cnpj}: {data}")
        else:
            # Lidar com erros
            print(f"Erro para CNPJ {cnpj}: {response.status_code} - {response.text}")
            
        
    
# Substitua "seu_token_aqui" pelo seu token real
token = "f2987ff033bd4550a7a208f4fc82be13c3ea145f2369561b089cba941742eda6"

todos_cnpj = ["01612795000151"]

buscar_cnpj_api(todos_cnpj, token)