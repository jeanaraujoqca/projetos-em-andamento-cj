import os
import pandas as pd
import time

BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
ARQUIVOS_DIR = os.path.join(BASE_DIR, 'arquivos')

# renomear os relatorios de acordo com o número do processo
print('Modificando os nomes dos arquivos de acordo com o número do processo')

df_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR) if file.endswith('.xlsx')][0]
df = pd.read_excel(df_path)

dicionario = {}
for indice, numero_processo in enumerate(df['Número Processo']):
    dicionario[indice] = numero_processo

# dicionario[int(os.path.splitext(os.listdir(ARQUIVOS_DIR)[0])[0].split('_')[-1])]

for relatorio in os.listdir(ARQUIVOS_DIR):
    if relatorio.endswith('.pdf'):
        numero = None
        nome_relatorio, extensao = os.path.splitext(relatorio)
        numero = int(nome_relatorio.split('_')[-1])

        # Verifica se o número existe no dicionário e renomeia o relatorio
        if numero is not None and numero in dicionario:
            novo_nome = rf"{dicionario[numero]}{extensao}".replace('/', '-')
            # Verifica se o novo nome já existe, se sim, adiciona um sufixo único
            while os.path.exists(os.path.join(ARQUIVOS_DIR, novo_nome)):
                novo_nome = rf"{dicionario[numero]}_{time.time()}{extensao}".replace('/', '-')
            os.rename(os.path.join(ARQUIVOS_DIR, relatorio), os.path.join(ARQUIVOS_DIR, novo_nome))
        elif numero is None:
            os.rename(os.path.join(ARQUIVOS_DIR, relatorio), os.path.join(ARQUIVOS_DIR, f"{dicionario[0]}{extensao}"))

print('\nO processo de renomeação foi finalizado.')
