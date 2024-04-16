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
import keyboard
import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import gspread_formatting
from gspread_formatting import Color
import gspread
import shutil
import xlrd
from openpyxl import Workbook
from io import StringIO
from features import utilities
from features import BrowserFunctions



# EXPORTAR PLANILHA HOJE

# Abrir navegador
driver = webdriver.Chrome()
BrowserFunctions.accessWebsite(driver)

# Fazer login
BrowserFunctions.login(driver)
time.sleep(6)

# Filtros para exportação
BrowserFunctions.openSearchPage(driver)
time.sleep(1)

hoje = datetime.now()
semana = hoje.weekday()

if semana == 0:
     ontem = hoje - timedelta(days=28)
     data_final = hoje
else:
     ontem = hoje - timedelta(days=29)
     data_final = hoje 
     
data = ontem.date().strftime("%d%m%Y")
data2 = data_final.date().strftime("%d%m%Y")

BrowserFunctions.exportExcelDelivered(driver, data, data2)

time.sleep(0.5)
driver.quit()

while not any(file.endswith('.xls') or file.endswith('.xlsx') for file in os.listdir(utilities.pasta_origem)):
    time.sleep(1)

# Mover o arquivo
utilities.moveFile(utilities.pasta_origem, utilities.pasta_destino, '.xls')
time.sleep(2)

# Renomear o arquivo
novo_nome = 'cargas_E.xls'
utilities.renameFile(utilities.pasta_destino, novo_nome, 1, '.xls')
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

# Lendo planilha cargas entregues
cargas_entregue_df = pd.read_excel('cargas_entregues.xlsx')

# Multiplica os valores por 0.01 para mover duas vírgulas para a esquerda
cargas_entregue_df['PESO'] = pd.to_numeric(
    cargas_entregue_df['PESO'], errors='coerce')
cargas_entregue_df['PESO'] = cargas_entregue_df['PESO'] * 0.01

# Filtrar por peso acima de 1kg
col_selecionadas2 = ["NFS/DOC-S"]
ind_coluna_especifica2 = 15
col_selecionadas2.append(cargas_entregue_df.columns[ind_coluna_especifica2])
cargas_refri_df2 = cargas_entregue_df.loc[cargas_entregue_df['PESO'] >= 1, col_selecionadas2]

 # Filtrando por Estado
indice_coluna_especifica3 = 1
caracteres_especificos2 = "- AP|- AM|- MS|- BA|- SE|- SC|- GO|- PA|- PI"  
coluna_especifica2 = cargas_refri_df2.columns[indice_coluna_especifica3]
linhas_filtradas_df2 = cargas_refri_df2[cargas_refri_df2[coluna_especifica2].str.contains(caracteres_especificos2)]

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
         
        # PESQUISAR NA PLANILHA AS LINHAS QUE CONTEM AS NFS JÁ ENTREGUES E SALVAR A INFORMAÇÃO DA LINHA
         spreadsheet_id = SAMPLE_SPREADSHEET_ID
         range_name = f"{SHEET_NAME}!D2:D"

         resultad = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
         nfs_doc_s_planilha = resultad.get('values', [])
         nfs_doc_s_planilha = [item for sublist in nfs_doc_s_planilha for item in sublist]
                  
         
         valores_comuns = list(set(nfs_doc_s_planilha) & set(nfs_doc_s_df_unique))
        

         # Selecionar os primeiros 60 valores
         primeiros_60 = valores_comuns[:60]

         # Selecionar os segundos 60 valores
         segundos_60 = valores_comuns[60:]

         print(f'Primeiros 60: ',primeiros_60)
         print(f'Segundos 60: ',segundos_60)

         
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
                     color = {"red": 0.5, "green": 0.5, "blue": 0.5}  # Cor cinza
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
                 
                     response = service.spreadsheets().batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
             time.sleep(60)
        
     except HttpError as err:
          print(err)


if __name__ == '__main__':
     main() 

time.sleep(0.5)

# Excluir arquivos temporários
utilities.deleteFiles(utilities.pasta_destino)
utilities.deleteFiles(utilities.pasta_destinoPDF)