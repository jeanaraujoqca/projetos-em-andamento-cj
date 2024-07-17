from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from datetime import date, timedelta
import pandas as pd
import os
from collections import defaultdict


# --- CARREGAR E TRATAR A BASE DE DADOS ---
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
ARQUIVOS_DIR = os.path.join(BASE_DIR, 'arquivos')

# carregar o arquivo de Substabelecimento.pdf
arquivo_substabelecimento = [os.path.join(BASE_DIR, file) for file in os.listdir(BASE_DIR) if file == 'Substabelecimento.pdf'][0]
print('Arquivo Substabelecimento.pdf carregado')

# Carrgar base de dados
df_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR) if file.endswith('.xlsx')][0]
nome_arquivo = df.split('\\')[-1]
df = pd.read_excel(df)
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
df = pd.merge(df, df_files_path, how='left', left_on='numero_processo', right_on='numero_processo')
df.columns
df.to_excel('ID_ARQUIVOS.xlsx', index=False)

# Selecionar apenas os processos que são Eletrônicos e que sao do responsavel
NOME_RESPONSAVEL = 'Provisório - Talita Mayara da Silva'
df = df[(df['Tipo de Processo'] == 'Eletrônico') & (df['Responsavel Prazo'] == str(NOME_RESPONSAVEL))]
print(df.head())
print(df.shape)

# --- INICIALIZAR O BROWSER --- [ X ]
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 20
wait = WebDriverWait(driver, timeout)


# ENTRAR NO PERFORMA [ X ]
url = 'https://performaqca.seven.adv.br/'
driver.get(url)
print('Espere o tempo do login e do carregamento da página')
print('Contando os 40 segundos')
sleep(40)
print('Entramos no Performa...')


# ENTRAR DIRETO NO ENDPOINT COM ID PRAZO --> Reaproveitar o código de Ana Luiza
# https://performaqca.seven.adv.br/compromisso/details/7794124
print('Iniciando preenchimento')

casos_sucesso = []
casos_fracasso = []
for indice, linha in df.iterrows():
    id_prazo = linha['ID Compromisso']
    num_processo = linha['numero_processo']
    endereco_arquivo = linha['arquivos']
    endpoint = f'https://performaqca.seven.adv.br/compromisso/details/{str(id_prazo)}'
    driver.get(endpoint)
    print(f'\n{indice+1}. Iniciando o processo: {num_processo} ({id_prazo})')
    sleep(5)
    try:
        
        # --- SCROOL PARA O ELEMENTO ---
        elemento_scroll = driver.find_element(By.XPATH, '//*[@id="DataProtocoloInterno"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(.5)
        
        # Colocar a data do prazo
        data_do_prazo = driver.find_element(By.XPATH, '//*[@id="DataProtocoloInterno"]')
        data_do_prazo.clear()
        data_do_prazo.send_keys(date.today().strftime('%d/%m/%Y'), Keys.ENTER)
        sleep(1)
        
        
        # ESCREVER NO CAMPO DE OBSERVAÇÃO
        observacao = driver.find_element(By.XPATH, '//*[@id="ObservacaoConclusao"]')
        observacao.send_keys('Peça incluída')
        sleep(1)
        
        # --- SCROOL PARA O ELEMENTO ---
        elemento_scroll = driver.find_element(By.XPATH, '//*[@id="ufile"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(.5)
        
        # ANEXO DOCS [ X ] --> Reaproveitar o código de anexo de mais de um arquivo da automação de Daiana
        anexos = driver.find_element(By.XPATH, '//*[@id="ufile"]')
        anexos.send_keys(rf'{endereco_arquivo}')
        anexos.send_keys(rf'{arquivo_substabelecimento}')
        sleep(1)
        
        # PARTE 1 ---------------------------------------------------------------------------------------------------------
        # Selecionar as opções necessárias
        # Clicar em "Tipo de Documento"
        tipo_documento = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[2]/div/button')
        tipo_documento.click()
        # Escrever o "Tipo de Documento"
        documento_input = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[2]/div/div/div[1]/input')
        documento_input.send_keys('Protocolo - Habilitação', Keys.ENTER)
        # Clicar em "Documento Interno = Sim"
        doc_interno_sim = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[2]/div/ins')[0]
        doc_interno_sim.click()
        # Clicar em "Foi utilizado Legal Design = Não"
        legal_design_nao = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[1]/div/ins')[-1]
        legal_design_nao.click()
        # Clicar em "Integrar protocolo? = Sim"
        integrar_protocolo_sim = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[2]/div/ins')[1]
        integrar_protocolo_sim.click()
        # Digitar a data atual no formato brasileiro
        data_atual = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[3]/div/input')
        data_atual.send_keys(date.today().strftime("%d/%m/%Y"), Keys.ENTER)
        sleep(1)

        # --- SCROOL PARA O ELEMENTO ---
        elemento_scroll = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(.5)

        # PARTE 2 ---------------------------------------------------------------------------------------------------------
        # Clicar em "Tipo de Documento"
        tipo_documento = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[2]/div/button')
        tipo_documento.click()
        # Escrever o "Tipo de Documento"
        documento_input = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[2]/div/div/div[1]/input')
        documento_input.send_keys('Protocolo - Habilitação', Keys.ENTER)
        # Clicar em "Documento Interno = Sim"
        doc_interno_sim = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[2]/div/ins')[0]
        doc_interno_sim.click()
        # Clicar em "Foi utilizado Legal Design = Não"
        legal_design_nao = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[1]/div/ins')[-1]
        legal_design_nao.click()
        # Clicar em "Integrar protocolo? = Sim"
        integrar_protocolo_sim = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[2]/div/ins')[1]
        integrar_protocolo_sim.click()
        # Digitar a data atual no formato brasileiro
        data_atual = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[3]/div/input')
        data_atual.send_keys(date.today().strftime("%d/%m/%Y"), Keys.ENTER)
        sleep(1)

        # CLICAR EM CONCLUIR [ X ]
        elemento_scroll = driver.find_element(By.XPATH, '//*[@id="buttonConcluir"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(0.5)
        
        botao_concluir = driver.find_element(By.XPATH, '//*[@id="buttonConcluir"]')
        botao_concluir.click()
        
        print(f'Processo: {num_processo} ({id_prazo}) preenchido')
        casos_sucesso.append({'Caso': id_prazo, 'Status': 'Sucesso'})
        sleep(5)
    except TimeoutException as e:
        print(f'{indice+1}: Erro ao processar o processo {id_prazo}')
        casos_fracasso.append({'Processo': id_prazo, 'Status': f'Erro: {str(e)}'})
    except Exception as e:
        print(f'{indice+1}: Erro inesperado ao processar o processo {id_prazo}')
        casos_fracasso.append({'Processo': id_prazo, 'Status': f'Erro inesperado: {str(e)}'})
    
# exportando as bases para controle do usuario
df_sucesso = pd.DataFrame(casos_sucesso)
df_fracasso = pd.DataFrame(casos_fracasso)
df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

sleep(1)

# fechando o browser para finalizar o procedimento
driver.quit()
print('\nProcesso finalizado.')

