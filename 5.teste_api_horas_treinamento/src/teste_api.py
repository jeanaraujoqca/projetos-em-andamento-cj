import requests

url = 'http://127.0.0.1:8000/run-rpa'

data = {
    'email': 'jeanaraujo@queirozcavalcanti.adv.br',
    'senha': 'Jlrc@2504'
}

files = {'file': open('C:/Users/jeanaraujo/ForaDoDrive/Operações/Automações/Andamento/3.automacao_lancamento_horas_treinamento_corrigida/data/Lançamento de Horas de Treinamentos(1-9) (4).xlsx', 'rb')}

response = requests.post(url, data=data, files=files)

print('Status Code:', response.status_code)
try:
    print('Response JSON:', response.json())
except ValueError:
    print('No JSON response')