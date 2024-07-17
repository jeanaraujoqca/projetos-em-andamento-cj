import pandas as pd
import os
from collections import defaultdict


# --- CARREGAR E TRATAR A BASE DE DADOS ---
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
ARQUIVOS_DIR = os.path.join(BASE_DIR, 'arquivos2')

# carregar o arquivo de Substabelecimento.pdf
arquivo_substabelecimento = [os.path.join(BASE_DIR, file) for file in os.listdir(BASE_DIR) if file == 'Substabelecimento.pdf'][0]
print('Arquivo Substabelecimento.pdf carregado')

# Carrgar base de dados
df_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR) if file.endswith('.xlsx')][0]
nome_arquivo = df_path.split('\\')[-1]
df = pd.read_excel(df_path)
print(f'Arquivo {nome_arquivo} carregado ')
print(f'Temos {df.shape} linhas')
print(df.columns)

# Listar todos os documentos da pasta de arquivos
files = os.listdir(ARQUIVOS_DIR)
files_path = [os.path.join(ARQUIVOS_DIR, file) for file in os.listdir(ARQUIVOS_DIR)]

# Criar um dataframe somente com os caminhos dos arquivos
df_files_path = pd.DataFrame(files_path, columns=['arquivos'])

# Procv para termos os caminhos de acordo com o ID Prazo
df_files_path['numero_processo'] = df_files_path['arquivos'].apply(lambda x: str(os.path.splitext(x)[0].split('\\')[-1]))
df = pd.merge(df, df_files_path, how='left', left_on='Número do Processo', right_on='numero_processo')
df.columns
df.to_excel('ID_ARQUIVOS2.xlsx', index=False)
