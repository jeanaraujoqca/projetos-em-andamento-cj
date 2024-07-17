from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import os


BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
df_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR) if file.endswith('.xlsx')][0]
df = pd.read_excel(df_path)

dicionario = {}
for processo in df['Número Processo']:
    if processo in dicionario:
        dicionario[processo] += 1
    else:
        dicionario[processo] = 1

# Inserir arquivo PDF
uploaded_pdf_file = r"C:\Users\victoramarante\Documents\automacao_rostand\PETICOES.pdf"
if os.path.exists(uploaded_pdf_file):
    # Criar uma pasta temporária para salvar os PDFs separados
    temp_dir = 'arquivos'
    if not os.path.exists(temp_dir):
        os.mkdir(temp_dir)

    pdf_reader = PdfReader(uploaded_pdf_file)
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
    
        # Criar um novo arquivo PDF para a página atual
        output_pdf = PdfWriter()
        output_pdf.add_page(page)

        # Salvar o novo arquivo PDF
        output_filename = os.path.join(temp_dir, f'output_{page_num}.pdf')
        with open(output_filename, 'wb') as f:
            output_pdf.write(f)
            
else:
    print('arquivo não foi encontrado')

