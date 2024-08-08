from playwright.sync_api import sync_playwright
from time import sleep
import pandas as pd
import numpy as np
import openpyxl
import os
import shutil
import json

# --- ORGANIZAR OS CAMINHOS DE DIRETÓRIO --- [ X ]
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, '5.atualizacao_follow_up_duda', 'data')
CONFIG_DIR = os.path.join(BASE_DIR, '5.atualizacao_follow_up_duda', 'config_senha')
arquivo_config = os.path.join(CONFIG_DIR, 'config_senha.json')

# --- SOLICITAR AS CREDENCIAIS PARA QUEM ESTÁ UTILIZANDO O PROGRAMA --- [ X ]
with open(arquivo_config, 'r') as config_file:
    config = json.load(config_file)
    nome_usuario = config['nome_usuario']
    email_usuario = config['email_usuario']
    senha_usuario = config['senha_usuario']
    
# --- CARREGAR A BASE DE DADOS COM AS INFORMAÇÕES DOS FOLLOW UPs --- [ X ]
file_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR)][0]
df_initial = pd.read_excel(file_path)
# df_initial = pd.read_excel(file_path)
# df = df_initial.copy()
df = df_initial

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto('https://performaqca.seven.adv.br/login')
    print(page.title())
    
    botao_confirmar = '//*[@id="idSIButton9"]'
    
    logar_credenciais_rede = '//a[@id="buttonVerificarUsuario"]'
    page.click(logar_credenciais_rede)
    
    user_email = '//*[@id="i0116"]'
    page.fill(user_email, email_usuario)
    page.click(botao_confirmar)
    sleep(2)
    
    user_password = '//*[@id="i0118"]'
    page.fill(user_password, senha_usuario)
    page.click(botao_confirmar)
    sleep(2)
    
    # botao_confirmar = '//*[@id="idSIButton9"]'
    page.click(botao_confirmar)
    sleep(10)
    
    print('Entramos no Performa. Seguindo para o passo seguinte...')
    print('Iniciando as iterações sobre os processos')
    
    for index, row in df.iterrows():
        processo = row['ID Processo']
        follow_up = row['Atualização']
        try:
            page.goto(f'https://performaqca.seven.adv.br/processo/details/{processo}')
            sleep(3)
            
            negociacao = '//ul//li//a[@href = "#box-negociacao"]'
            page.click(negociacao)
            
            status_acordo = '//*[@id="negociacaoList"]/div[1]/table/tbody/tr[1]/td[2]/span'
            status_text = page.inner_text(status_acordo)
            
            responsavel_acordo = '//*[@id="negociacaoList"]/div[1]/table/tbody/tr[8]/td[2]/span'
            responsavel_text = page.inner_text(responsavel_acordo)
            
            if (status_text == 'Em Negociação' or status_text == 'Em Negociação Pós Sentença') and (responsavel_text == nome_usuario):
                print(f"O processo está liberado para inserir o Follow Up de Acordo.")
                sleep(1)
                
                botao_follow_up = '//*[@id="buttonNovoFollowupAcordo"]'
                page.click(botao_follow_up)
                sleep(1)
                
                adv_parte = '//ins[@class="iCheck-helper"]'
                page.click(adv_parte[1])
                
                observacao = '//*[@id="Observacao"]'
                observacao.fill(follow_up)
                sleep(1)
                
                botao_salvar = '//*[@id="dialog-modal"]/div/div/div[3]/button[1]'
                page.click(botao_salvar)
                print(f'Processo {index+1}: {processo}\n')
                sleep(3)
            else:
                print('Bloqueado')   
        except Exception as e:
            print(f'Erro ao processar o processo {processo}\n')
            continue
        
    browser.close()

print('O processo foi finalizado.')