import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from time import sleep
import os

# Function to handle form submission
def submit_form():
    # Retrieve user inputs
    email = email_entry.get()
    senha = senha_entry.get()
    file_path = file_path_label.cget("text")
    
    # Check if email is provided
    if not email:
        messagebox.showwarning("Aviso", "Por favor, insira o e-mail.")
        return
    
    if not senha:
        messagebox.showwarning("Aviso", "Por favor, insira a senha.")
        return
    
    # Load DataFrame from file
    file_path = file_path_label['text']
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)

        # Initialize Selenium
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(service=service, options=options)
        driver.maximize_window()
        timeout = 15
        wait = WebDriverWait(driver, timeout)

        # SharePoint URL
        url_sharepoint = 'https://queirozcavalcanti.sharepoint.com/sites/qca360/Lists/treinamentos_qca/AllItems.aspx'

        try:
            # Navigate to SharePoint
            driver.get(url_sharepoint)
            sleep(5)

            # Handle login
            try:
                email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='i0116']")))
                email_input.send_keys(email)
                sleep(.5)
                avancar_button = driver.find_element(By.XPATH, "//input[@id='idSIButton9']")
                avancar_button.click()
            except:
                pass

            try:
                senha_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='i0118']")))
                senha_input.send_keys(senha)
                sleep(.5)
                avancar_button = driver.find_element(By.XPATH, "//input[@id='idSIButton9']")
                avancar_button.click()
                sleep(3)
            except:
                pass

            try:
                botao_sim = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='idSIButton9']")))
                botao_sim.click()
            except:
                pass

            print('Entramos no Sharepoint. Aguarde para iniciar o procedimento de preenchimento das informações dos treinamentos.')
            sleep(10)

            print('Iniciando preenchimento')

            casos_sucesso = []
            casos_fracasso = []

            # Iterate over DataFrame rows
            for index, id in enumerate(df['ID']):
                print('Entramos dentro do laço de repetição')
                try:
                    colaborador = df.loc[index, 'Nome']
                    email_colaborador = df.loc[index, 'Email']
                    equipe = df.loc[index, 'EQUIPE']
                    unidade = df.loc[index, 'UNIDADE']
                    treinamento = df.loc[index, 'TREINAMENTO']
                    tipo_de_treinamento = df.loc[index, 'TIPO DO TREINAMENTO']
                    categoria = df.loc[index, 'CATEGORIA']
                    instituicao_instrutor = df.loc[index, 'INSTITUIÇÃO/INSTRUTOR']
                    carga_horaria = df.loc[index, 'CARGA HORÁRIA']
                    inicio_do_treinamento = df.loc[index, 'INICIO DO TREINAMENTO']
                    termino_do_treinamento = df.loc[index, 'TERMINO DO TREINAMENTO']

                    print(f'Adicionando informações do colaborador: {colaborador}')

                    # Add new training
                    botao_novo = driver.find_element(By.XPATH, '//*[@id="appRoot"]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[1]/button/span')
                    botao_novo.click()                          
                    sleep(5)
                    # cuidado com esse sleep pode ter ocasionado alguma bronca
                    # Switch to iframe
                    iframe = driver.find_elements(By.XPATH, "//iframe")
                    driver.switch_to.frame(iframe[0])

                    iframe2 = driver.find_element(By.XPATH, "//iframe[@class='player-app-frame']")
                    driver.switch_to.frame(iframe2)
                    sleep(3)

                    # Function to interact with elements
                    def clica_seleciona_informacao(endereco1, endereco2, valor2, endereco3):
                        elemento1 = wait.until(EC.element_to_be_clickable((By.XPATH, endereco1)))
                        elemento1.click()
                        sleep(3)
                        elemento2 = wait.until(EC.element_to_be_clickable((By.XPATH, endereco2)))
                        elemento2.send_keys(valor2)
                        sleep(3)
                        elemento3 = wait.until(EC.element_to_be_clickable((By.XPATH, endereco3)))
                        elemento3.click()

                    # Fill form fields
                    clica_seleciona_informacao(endereco1='//div[@title="NOME DO INTEGRANTE"]',
                                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-0"]/div/div/div/div/input', valor2=str(colaborador),
                                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-0"]/div/ul/li/div/div/span[1][text() = "{str(colaborador)}"]')
                    sleep(2)

                    clica_seleciona_informacao(endereco1='//div[@title="E-MAIL"]',
                                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-1"]/div/div/div/div/input', valor2={email_colaborador},
                                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-1"]/div/ul/li/div/span[text() = "{str(email_colaborador)}"]')
                    sleep(2)

                    clica_seleciona_informacao(endereco1='//div[@title="EQUIPE."]',
                                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-2"]/div/div/div/div/input', valor2=str(equipe),
                                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-2"]/div/ul/li/div/span[text() = "{str(equipe)}"]')
                    sleep(2)

                    clica_seleciona_informacao(endereco1='//div[@title="UNIDADE"]',
                                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-3"]/div/div/div/div/input', valor2=str(unidade),
                                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-3"]/div/ul/li/div/span[text() = "{str(unidade)}"]')
                    sleep(2)

                    driver.find_element(By.XPATH, '//input[@title="TREINAMENTO"]').send_keys(str(treinamento))
                    sleep(2)

                    clica_seleciona_informacao(endereco1='//div[@title="TIPO DO TREINAMENTO."]',
                                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-4"]/div/div/div/div/input', valor2=str(tipo_de_treinamento),
                                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-4"]/div/ul/li/div/span[text() = "{str(tipo_de_treinamento)}"]')
                    sleep(2)

                    elemento_scroll = driver.find_element(By.XPATH, '//input[@title="INSTITUIÇÃO/INSTRUTOR"]')
                    driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
                    sleep(.5)

                    clica_seleciona_informacao(endereco1='//div[@title="CATEGORIA"]',
                                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-5"]/div/div/div/div/input', valor2=str(categoria),
                                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-5"]/div/ul/li/div/span[text() = "{str(categoria)}"]')
                    sleep(2)

                    driver.find_element(By.XPATH, '//input[@title="INSTITUIÇÃO/INSTRUTOR"]').send_keys(str(instituicao_instrutor))
                    sleep(2)

                    clica_seleciona_informacao(endereco1='//div[@title="CARGA HORARIA"]',
                                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-6"]/div/div/div/div/input', valor2=str(carga_horaria),
                                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-6"]/div/ul/li/div/span[text() = "{str(carga_horaria)}"]')
                    sleep(2)

                    data_inicio = driver.find_element(By.XPATH, '//input[@title="INICIO DO TREINAMENTO"]')
                    data_inicio.send_keys(str(inicio_do_treinamento))
                    sleep(2)

                    data_final = driver.find_element(By.XPATH, '//input[@title="TERMINO DO TREINAMENTO"]')
                    data_final.send_keys(str(termino_do_treinamento))
                    sleep(2)

                    driver.switch_to.default_content()
                    salvar = driver.find_element(By.XPATH, '//span[text() = "Salvar"]')
                    salvar.click()

                    print(f'{index+1} - ID {id} - {colaborador} - {treinamento} - finalizado')

                    casos_sucesso.append({'Caso': id, 'Status': 'Sucesso'})
                    sleep(3)
                except TimeoutException as e:
                    print(f'Erro ao processar o treinamento {id}')
                    casos_fracasso.append({'Treinamento': id, 'Status': f'Erro: {str(e)}'})

                except Exception as e:
                    print(f'Erro inesperado ao processar o treinamento {id}')
                    casos_fracasso.append({'Treinamento': id, 'Status': f'Erro inesperado: {str(e)}'})

            df_sucesso = pd.DataFrame(casos_sucesso)
            df_fracasso = pd.DataFrame(casos_fracasso)
            df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
            df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

            sleep(2)

            driver.quit()

            print('O processo finalizou! Verificar no Sharepoint as informações editadas.')

        except Exception as e:
            print(f'Erro ao iniciar o navegador: {str(e)}')
            driver.quit()

# Function to handle file selection
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    file_path_label.config(text=file_path)
    if file_path:
        submit_button.pack()
        messagebox.showinfo("Alerta", "Alerta de duas etapas")  

# Initialize Tkinter
root = tk.Tk()
root.title('Automatização de Preenchimento - SharePoint')

# Criar elementos da GUI
email_label = tk.Label(root, text='Email:')
email_label.pack()
email_entry = tk.Entry(root)
email_entry.pack()

senha_label = tk.Label(root, text='Senha:')
senha_label.pack()
senha_entry = tk.Entry(root, show='*')
senha_entry.pack()

file_label = tk.Label(root, text='Selecione o arquivo Excel:')
file_label.pack()
file_path_label = tk.Label(root, text='')
file_path_label.pack()

file_button = tk.Button(root, text='Selecionar Arquivo', command=select_file)
file_button.pack()

submit_button = tk.Button(root, text='Iniciar Preenchimento', command=submit_form)

root.mainloop()