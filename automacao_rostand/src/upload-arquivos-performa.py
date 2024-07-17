from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from datetime import date
import pandas as pd
import os
from tqdm import tqdm

# --- CARREGAR E TRATAR A BASE DE DADOS ---
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')

# carregar o arquivo de Substabelecimento.pdf
arquivo_substabelecimento = [os.path.join(BASE_DIR, file) for file in os.listdir(BASE_DIR) if file == 'Substabelecimento.pdf'][0]
print('Arquivo Substabelecimento.pdf carregado')

# Carrgar base de dados
df_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR) if file.endswith('.xlsx')][-1]
nome_arquivo = df_path.split('\\')[-1]
df = pd.read_excel(df_path)
print(f'Arquivo {nome_arquivo} carregado ')
print(f'Temos {df.shape} linhas')

# --- INICIALIZAR O BROWSER ---
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 20
wait = WebDriverWait(driver, timeout)

# ENTRAR NO PERFORMA
url = 'https://performaqca.seven.adv.br/'
driver.get(url)
print('Espere o tempo do login e do carregamento da página')
print('Contando os 40 segundos')
sleep(40)
print('Entramos no Performa...')

# ENTRAR DIRETO NO ENDPOINT COM ID PRAZO
print('Iniciando preenchimento')

casos_sucesso = []
casos_fracasso = []

for indice, linha in tqdm(df.iterrows(), total=df.shape[0], desc="Processando"):
    id_prazo = linha['ID Prazo']
    num_processo = linha['numero_processo']
    endereco_arquivo = linha['arquivos']
    endpoint = f'https://performaqca.seven.adv.br/compromisso/details/{str(id_prazo)}'
    driver.get(endpoint)
    print(f'\n{indice+1}. Iniciando o processo: {num_processo} ({id_prazo})')
    sleep(5)
    try:
        # SCROOL PARA O ELEMENTO
        elemento_scroll = driver.find_element(By.XPATH, '//*[@id="DataProtocoloInterno"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(0.5)

        # Colocar a data do prazo
        data_do_prazo = driver.find_element(By.XPATH, '//*[@id="DataProtocoloInterno"]')
        data_do_prazo.clear()
        data_do_prazo.send_keys(date.today().strftime('%d/%m/%Y'), Keys.ENTER)
        sleep(1)

        # ESCREVER NO CAMPO DE OBSERVAÇÃO
        observacao = driver.find_element(By.XPATH, '//*[@id="ObservacaoConclusao"]')
        observacao.send_keys('Peça incluída')
        sleep(1)

        # SCROOL PARA O ELEMENTO
        elemento_scroll = driver.find_element(By.XPATH, '//*[@id="ufile"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(0.5)

        # ANEXO DOCS
        anexos = driver.find_element(By.XPATH, '//*[@id="ufile"]')
        anexos.send_keys(rf'{endereco_arquivo}')
        anexos.send_keys(rf'{arquivo_substabelecimento}')
        sleep(1)

        # PARTE 1
        tipo_documento = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[2]/div/button')
        tipo_documento.click()
        sleep(.3)
        documento_input = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[2]/div/div/div[1]/input')
        documento_input.send_keys('Protocolo - Habilitação', Keys.ENTER)
        doc_interno_sim = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[2]/div/ins')[0]
        doc_interno_sim.click()
        legal_design_nao = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[1]/div/ins')[-1]
        legal_design_nao.click()
        integrar_protocolo_sim = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[2]/div/ins')[1]
        integrar_protocolo_sim.click()
        data_atual = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[3]/div/input')
        data_atual.send_keys(date.today().strftime("%d/%m/%Y"), Keys.ENTER)
        sleep(1)

        elemento_scroll = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(0.5)

        # PARTE 2
        tipo_documento = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[2]/div/button')
        tipo_documento.click()
        sleep(.3)
        documento_input = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[2]/div/div/div[1]/input')
        documento_input.send_keys('Protocolo - Habilitação', Keys.ENTER)
        doc_interno_sim = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[2]/div/ins')[0]
        doc_interno_sim.click()
        legal_design_nao = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[1]/div/ins')[-1]
        legal_design_nao.click()
        integrar_protocolo_sim = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[2]/div/ins')[1]
        integrar_protocolo_sim.click()
        data_atual = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[3]/div/input')
        data_atual.send_keys(date.today().strftime("%d/%m/%Y"), Keys.ENTER)
        sleep(1)

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
