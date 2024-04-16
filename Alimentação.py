from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import googleapiclient.discovery
import googleapiclient.errors
import pyautogui
from google.oauth2 import service_account
from datetime import datetime, timedelta
from time import sleep
import time
import pandas as pd
import os
import PyPDF2
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import gspread_formatting
from gspread_formatting import Color
import gspread
from io import StringIO
from features import utilities
from features import BrowserFunctions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import numpy as np


# Abrir navegador
driver = webdriver.Chrome()
BrowserFunctions.accessWebsite(driver)

# Fazer login
BrowserFunctions.login(driver)
time.sleep(6)


# EXPORTAR PLANILHA NÃO ENTREGUES
BrowserFunctions.openSearchPage(driver)
time.sleep(1)

hoje = datetime.now()
semana = hoje.weekday()

if semana == 0:
     ontem = hoje - timedelta(days=3)
     data_final = hoje - timedelta(days=1)
else:
     ontem = hoje - timedelta(days=1)
     data_final = hoje - timedelta(days=1)

data = ontem.date().strftime("%d%m%Y")
data2 = data_final.date().strftime("%d%m%Y")

BrowserFunctions.exportExcelUndelivered(driver, data, data2)

while not any(file.endswith('.xls') or file.endswith('.xlsx') for file in os.listdir(utilities.pasta_origem)):
    time.sleep(1)

# Mover o arquivo
utilities.moveFile(utilities.pasta_origem, utilities.pasta_destino, '.xls')
time.sleep(2)

# Renomear o arquivo
novo_nome = 'cargas_SR.xls'
utilities.renameFile(utilities.pasta_destino, novo_nome, 1, '.xls')
time.sleep(4)

# Lê o conteúdo do arquivo 'cargas_sem_recebedor.xls' para uma string
with open('cargas_SR.xls', 'r', encoding='utf-8') as file:
    html_content = file.read()
# Usa StringIO para criar um objeto de leitura da string
html_io = StringIO(html_content)
# Lê as tabelas do HTML usando o objeto StringIO
tables = pd.read_html(html_io)
# Agora você pode acessar as tabelas como DataFrames
for idx, table in enumerate(tables):
    # Salva a tabela como arquivo XLSX
    table.to_excel('cargas_sem_recebedor.xlsx', index=False)

time.sleep(3)

# EXPORTAR PLANILHA ENTREGUES

# Filtros para exportação
BrowserFunctions.openSearchPage(driver)
time.sleep(1)

if semana == 0:
    ontem = hoje - timedelta(days=28)
    data_final = hoje
else:
    ontem = hoje - timedelta(days=29)
    data_final = hoje

data = ontem.date().strftime("%d%m%Y")
data2 = data_final.date().strftime("%d%m%Y")

BrowserFunctions.exportExcelDelivered(driver, data, data2)

while not any(file.endswith('.xls') or file.endswith('.xlsx') for file in os.listdir(utilities.pasta_origem)):
    time.sleep(1)

# Mover o arquivo
utilities.moveFile(utilities.pasta_origem, utilities.pasta_destino, '.xls')
time.sleep(2)

# Renomear o arquivo
novo_nome = 'cargas_E.xls'
utilities.renameFile(utilities.pasta_destino, novo_nome, 2, '.xls')
time.sleep(4)

# Lê o conteúdo do arquivo 'cargas_sem_recebedor.xls' para uma string
with open('cargas_E.xls', 'r', encoding='utf-8') as file:
    html_content = file.read()
# Usa StringIO para criar um objeto de leitura da string
html_io = StringIO(html_content)
# Lê as tabelas do HTML usando o objeto StringIO
tables = pd.read_html(html_io)
# Agora você pode acessar as tabelas como DataFrames
for idx, table in enumerate(tables):
    # Salva a tabela como arquivo XLSX
    table.to_excel('cargas_entregues.xlsx', index=False)
time.sleep(4)

# Extrair planilha Cliente

driver.get('Endereço de E-mail')

# Fazer Login
campo_usuario2 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rcmloginuser"]')))
campo_usuario2.send_keys('Usuário')
campo_senha2 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rcmloginpwd"]')))
campo_senha2.send_keys('senha')
botao_entrar2 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rcmloginsubmit"]')))
botao_entrar2.click()
time.sleep(3)

pesquisa_email = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mailsearchform"]')))
pesquisa_email.send_keys('planilha de conservação')
time.sleep(1)
pesquisa_email.send_keys(Keys.ENTER)
time.sleep(2)
pyautogui.click(476,296, duration=2.5)
pyautogui.click(883,325, duration=2.5)
pyautogui.click(775,163, duration=2.5)
pyautogui.click(1199,36, duration=2)

while not any(file.endswith('.xls') or file.endswith('.xlsx') for file in os.listdir(utilities.pasta_origem)):
    time.sleep(1)

# Mover arquivo
utilities.moveFile(utilities.pasta_origem, utilities.pasta_destino, '.xlsx')

time.sleep(2)

# Renomear o arquivo
novo_nome = 'Planilha de Conservação Luty Log.xlsx'
utilities.renameFile(utilities.pasta_destino, novo_nome, 3, '.xlsx')

time.sleep(1)


#  Lendo a planilha cargas ontem
cargas_SR_df = pd.read_excel("cargas_sem_recebedor.xlsx")


# Multiplica os valores por 0.01 para mover duas vírgulas para a esquerda
cargas_SR_df['PESO'] = pd.to_numeric(cargas_SR_df['PESO'], errors='coerce')
cargas_SR_df['PESO'] = cargas_SR_df['PESO'] * 0.01

# Filtrar por peso acima de 1kg

colunas_selecionadas = ["AWB", "DATA COLETA", "NFS/DOC-S"]
indice_coluna_especifica = 15
coluna_especifica = cargas_SR_df.columns[indice_coluna_especifica]
colunas_selecionadas.append(coluna_especifica)
cargas_refri_df = cargas_SR_df.loc[cargas_SR_df['PESO']
                                   >= 1, colunas_selecionadas]

# Filtrando por Estado
indice_coluna_especifica = 3
caracteres_especificos = "- AP|- AM|- MS|- BA|- SE|- SC|- GO|- PA|- PI"
coluna_especifica = cargas_refri_df.columns[indice_coluna_especifica]
linhas_filtradas_df = cargas_refri_df[cargas_refri_df[coluna_especifica].str.contains(
    caracteres_especificos)]

linhas_filtradas_df['NFS/DOC-S'] = linhas_filtradas_df['NFS/DOC-S'].astype(str)

# Dividir os valores da coluna específica e replicar os dados das outras colunas
new_rows = []
for index, row in linhas_filtradas_df.iterrows():
    values = row['NFS/DOC-S'].split(',')
    for value in values:
        new_row = row.copy()
        new_row['NFS/DOC-S'] = value
        new_rows.append(new_row)

new_df = pd.DataFrame(new_rows)

new_df['NFS/DOC-S'] = new_df['NFS/DOC-S'].astype(float)

# Lendo planilha conservação Luty
conservacao_df = pd.read_excel("Planilha de Conservação Luty Log.xlsx")
conservacao_df.rename(columns={'Nota Fiscal': 'NFS/DOC-S'}, inplace=True)
# nomeAntigo = conservacao_df.columns[3]
if 'Embalagem' in conservacao_df.columns:
    print('OK')
elif 'Embalagem ' in conservacao_df.columns:
    conservacao_df.rename(columns={'Embalagem ': 'Embalagem'}, inplace=True)
    print('ESPAÇO')
elif 'CAIXA' in conservacao_df.columns:
    conservacao_df.rename(columns={'CAIXA': 'Embalagem'}, inplace=True)
    print('CAIXA')
else:
    conservacao_df['Embalagem'] = 'CAIXA EPS 12 LTS PS S/ GELO'
    print('CRIADO COLUNA')                           

# Mesclar planilhas
merged_df = pd.merge(new_df, conservacao_df, on='NFS/DOC-S', how='left')
coluna_comum = 'NFS/DOC-S'
merged_df[coluna_comum] = merged_df.apply(lambda row: row[coluna_comum] if not pd.isna(
    row['NFS/DOC-S']) else row[coluna_comum], axis=1)

# Retirando os não refrigerados
coluna_verificar = 'Possui Refrigerado?'
dado_especifico = 'N'
merged_df = merged_df[merged_df[coluna_verificar] != dado_especifico]


dados_update = merged_df

# Prepara os dados

dados_update['AWB'] = dados_update['AWB'].fillna("")
dados_update['Embalagem'] = dados_update['Embalagem'].fillna("")
dados_update['INFO ADICIONAL'] = dados_update.apply(
    lambda row: '72h' if (row['Embalagem'] == 'CAIXA EPS 12 LTS PD C/ GELO') or (row['Embalagem'] == 'CAIXA EPS 44 LTS PD C/ GELO') else '48h', axis=1)
dados_update['PROX MANUT'] = pd.to_datetime(
    dados_update['DATA COLETA'], format='%d/%m/%Y', errors='coerce')
dados_update.loc[dados_update['INFO ADICIONAL'] ==
                 '48h', 'PROX MANUT'] += pd.Timedelta(days=2)
dados_update.loc[dados_update['INFO ADICIONAL'] ==
                 '72h', 'PROX MANUT'] += pd.Timedelta(days=3)
dados_update['PROX MANUT'] = dados_update['PROX MANUT'].fillna("").astype(str)
dados_update['PROX MANUT'] = pd.to_datetime(
    dados_update['PROX MANUT'], errors='coerce').dt.strftime('%d/%m/%Y')

dados_update['DATA COLETA'] = dados_update['DATA COLETA'].fillna(
    "").astype(str)
dados_update['DATA COLETA'] = pd.to_datetime(
    dados_update['DATA COLETA'], errors='coerce').dt.strftime('%d/%m/%Y')

dados_update['EMPRESA'] = ''
dados_update.loc[dados_update['Embalagem'].isin(
    ['CAIXA EPS 12 LTS PD C/ GELO', 'CAIXA EPS 12 LTS PS C/ GELO', 'CAIXA EPS 44 LTS PS C/ GELO', 'CAIXA EPS 44 LTS PD C/ GELO']), 'EMPRESA'] = '4BIO'

dados_update['NFS/DOC-S'] = dados_update['NFS/DOC-S'].replace([np.nan, np.inf, -np.inf], 0)

dados_update['NFS/DOC-S'] = dados_update['NFS/DOC-S'].astype(int)

colunas_escolhidas = ['AWB', 'DATA COLETA', 'NFS/DOC-S',
                               'CIDADE.1', 'INFO ADICIONAL', 'PROX MANUT', 'EMPRESA']

dados_update = dados_update[colunas_escolhidas]

print(dados_update)

BrowserFunctions.accessWebsite(driver)
BrowserFunctions.login(driver)

time.sleep(3)

for NFE in dados_update['NFS/DOC-S']:
    nomeArquivo = f'{NFE}.pdf'
    filePath = os.path.join(utilities.pasta_destinoPDF, nomeArquivo)
    BrowserFunctions.openSearchPage(driver)
    time.sleep(1.5)
    refrigeratedMedicine = BrowserFunctions.downloadNFE(driver,NFE, nomeArquivo)
    print(refrigeratedMedicine)
    if refrigeratedMedicine == False:
        dados_update = dados_update[dados_update['NFS/DOC-S'] != NFE]
    elif refrigeratedMedicine == True:
        company = utilities.filterCompany(filePath)
        dados_update.loc[dados_update['NFS/DOC-S'] == NFE, 'EMPRESA'] = company
    else:
        linesToUpdate = dados_update[(dados_update['NFS/DOC-S'] == NFE) & (dados_update['EMPRESA'] == '')]
        if not linesToUpdate.empty:
            dados_update.loc[dados_update['NFS/DOC-S'] == NFE, 'EMPRESA'] = refrigeratedMedicine
       
print(dados_update)
    
    
driver.quit()

# Lendo planilha cargas entregues
cargas_entregue_df = pd.read_excel("cargas_entregues.xlsx")

# Multiplica os valores por 0.01 para mover duas vírgulas para a esquerda
cargas_entregue_df['PESO'] = pd.to_numeric(
    cargas_entregue_df['PESO'], errors='coerce')
cargas_entregue_df['PESO'] = cargas_entregue_df['PESO'] * 0.01

# Filtrar por peso acima de 1kg
col_selecionadas2 = ["NFS/DOC-S"]
ind_coluna_especifica2 = 15
col_selecionadas2.append(cargas_entregue_df.columns[ind_coluna_especifica2])
cargas_refri_df2 = cargas_entregue_df.loc[cargas_entregue_df['PESO']
                                          >= 1, col_selecionadas2]

# Filtrando por Estado
indice_coluna_especifica3 = 1
caracteres_especificos2 = "- AP|- AM|- MS|- BA|- SE|- SC|- GO|- PA|- PI"
coluna_especifica2 = cargas_refri_df2.columns[indice_coluna_especifica3]
linhas_filtradas_df2 = cargas_refri_df2[cargas_refri_df2[coluna_especifica2].str.contains(
    caracteres_especificos2)]


nfs_doc_s_df_unique = linhas_filtradas_df2.iloc[:, 0].unique()
nfs_doc_s_df_unique = [str(item).strip() for item in nfs_doc_s_df_unique]


# Adiciona todos os escopos ao acesso à planilha
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


# O ID e o intervalo de uma planilha de amostra.
SAMPLE_SPREADSHEET_ID = 'ID da planilha google sheets'
SHEET_NAME = 'CONTROLE'


# Fazer login no google sheets
def main():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Alterações na planilha
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        # Busca qual a próxima linha vazia para preenchimento
        coluna = "C"

        resultado = service.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                                        range=f'CONTROLE!{coluna}:{coluna}').execute()
        val = resultado.get('values', [])
        proxima = len(val) + 1
        proxima_linha = f"B{proxima}"
        print(proxima_linha)

        # Alimenta os dados
        
        valores_adicionar = dados_update[colunas_escolhidas].values.tolist()
        print(valores_adicionar)

        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                       range=proxima_linha, valueInputOption="USER_ENTERED", body={'values': valores_adicionar}).execute()

        # PESQUISAR NA PLANILHA AS LINHAS QUE CONTEM AS NFS JÁ ENTREGUES E SALVAR A INFORMAÇÃO DA LINHA
        spreadsheet_id = SAMPLE_SPREADSHEET_ID
        range_name = f"{SHEET_NAME}!D2:D"

        resultad = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_name).execute()
        nfs_doc_s_planilha = resultad.get('values', [])
        nfs_doc_s_planilha = [
            item for sublist in nfs_doc_s_planilha for item in sublist]

        valores_comuns = list(set(nfs_doc_s_planilha) &
                              set(nfs_doc_s_df_unique))

        # Selecionar os primeiros 60 valores
        primeiros_60 = valores_comuns[:60]

        # Selecionar os segundos 60 valores
        segundos_60 = valores_comuns[60:]

        print(f"primeiros 60: ", primeiros_60)
        print(f"segundos 60: ", segundos_60)

        sheet_id = 0

        for valores in [primeiros_60, segundos_60]:
            for valor_comum in valores:
                # Encontre o índice da linha que contém o valor_comum na coluna NFS/DOC-S
                linha_encontrada = None
                for index, valor_planilha in enumerate(nfs_doc_s_planilha, start=2):
                    if valor_planilha == valor_comum:
                        linha_encontrada = index
                        break

                if linha_encontrada:
                    # Defina a faixa de colunas a ser formatada
                    range_notation = f'A{linha_encontrada}:Z{linha_encontrada}'
                    color = {"red": 0.5, "green": 0.5,
                             "blue": 0.5}  # Cor cinza
                    body = {
                        "requests": [
                            {
                                "repeatCell": {
                                    "range": {"sheetId": sheet_id, "startRowIndex": linha_encontrada - 1, "endRowIndex": linha_encontrada},
                                    "cell": {"userEnteredFormat": {"backgroundColor": color}},
                                    "fields": "userEnteredFormat.backgroundColor",
                                }
                            }
                        ]
                    }

                    response = service.spreadsheets().batchUpdate(
                        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
            time.sleep(60)

    except HttpError as err:
        print(err)

if __name__ == '__main__':
    main()
time.sleep(0.5)

# Excluir arquivos temporários
utilities.deleteFiles(utilities.pasta_destino)
utilities.deleteFiles(utilities.pasta_destinoPDF)