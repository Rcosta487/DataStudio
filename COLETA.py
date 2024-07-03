
from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import socket
socket.setdefaulttimeout(30 * 120)

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

import pyodbc
import pandas as pd
#import win32com.client as win32

import time
from datetime import datetime
from datetime import date

HOJESTR=(datetime.today().strftime('%d/%m/%Y'))
HOJE = datetime.strptime(HOJESTR, '%d/%m/%Y').date()


server = 'xxxxxx' 
database = 'xx' 
username = 'xx' 
password = 'xxxxxxxx' 
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

    
def ESTOQUE_PROF_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1BAWRqVJBBBtm-Odjv_jcrk0L7SnKl1PwYm3xQC2u23w'
    SAMPLE_RANGE_NAME = 'Página1!A1:N589468'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_ESTOQUE_PROF}).execute()
def ESTOQUE_PROF_PARC_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1mJe74TbOdg1cGPYlYmrKaRSj8OK8ESN6Ot5IcT1WF3E'
    SAMPLE_RANGE_NAME = 'Página1!A1:N549000'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_ESTOQUE_PROF_PARC}).execute()

def ESTOQUE_SOC_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1VS2ZlqBgTIy75Q1AIy3NbVM1Pz01jUwo1k_BbZ1B0Zk'
    SAMPLE_RANGE_NAME = 'Página1!A1:N579468'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_ESTOQUE_SOC}).execute()
    
   
    
def ESTOQUE_PFK_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1Z1hocIdiC-N1DiLOks0pxMu28Td--1NOkADVwdnS2AA'
    SAMPLE_RANGE_NAME = 'Página1!A1:N579468'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_ESTOQUE_PFK}).execute()

def ESTOQUE_PJK_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1bcTFc7TEBIfgQpwYZY8As7zepKB-FYxdJBWUX4-fC_k'
    SAMPLE_RANGE_NAME = 'Página1!A1:N579468'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_ESTOQUE_PJK}).execute()
    
def PAGAMENTOS_sheets():
    
    SAMPLE_SPREADSHEET_ID = '16t_siG-OXZEd6yMqXWYW4izRp2md5aZfStnyDMSUz04'
    SAMPLE_RANGE_NAME = 'Página1!A1:B1000001'
    clear_values_request_body ={}
    creds = None
      
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    service = build('sheets', 'v4', credentials=creds)
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_PAGAMENTOS}).execute()    
    
    
def RESULT_NOTIFICAO_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1cEZjxn4ZJw2ne4qzNgrQt-h3NK62P56Ld3gmbjPS1-I'
    SAMPLE_RANGE_NAME = 'Página1!A1:D101000'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_RESULT_NOTIFICACAO}).execute()

def PERFIL_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1L3YEZLm0KMzW_xfSeRqUndxBeJ9OjYztCCtEy9Qtn5k'
    SAMPLE_RANGE_NAME = 'Página1!A1:H101000'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_PERFIL}).execute() 

def PERFIL_SOC_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1ZBJ-ruQSztH733vf457HEtvVbNzGAZ8gWSt9Ngo-zN4'
    SAMPLE_RANGE_NAME = 'Página1!A1:H101000'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_PERFIL_SOC}).execute() 

def TOTAL_ATIVOS_PROF_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1fM-JioaQi4Xln-KvpmXwJGVyTqhOQJQpWEoGB4HNmm4'
    SAMPLE_RANGE_NAME = 'Página1!A1:B2'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_TOTAL_ATIVOS}).execute()   

def TOTAL_ATIVOS_SOC_sheets():
    
    SAMPLE_SPREADSHEET_ID = '10ywNg8DWopyXmCGTX4dY-7MWb4SLxkTyJba5gOTnZb8'
    SAMPLE_RANGE_NAME = 'Página1!A1:B2'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_TOTAL_SOC_ATIVOS}).execute()       
    
def SOCIEDADES_POR_CIDADE_PAGO_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1cgNbGx2PKTs1JBGFUdNz7ybTn7XZFI5MvlDaQjn0OyE'
    SAMPLE_RANGE_NAME = 'Página1!A1:C11000'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'wb') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_SOCIEDADES_POR_CIDADE_PAGO}).execute()   
    
def SOCIEDADES_POR_CIDADE_NAOPAGO_sheets():
    
    SAMPLE_SPREADSHEET_ID = '1-_TmAMhp2cz7bg47oxUIxOOhKPDWdKOcIpdPZUT-6mA'
    SAMPLE_RANGE_NAME = 'Página1!A1:C10000'
    clear_values_request_body ={}
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               r'E:\Meu Drive\_ARQUIVOS PESSOAIS\POWER BI\client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    
    
    request = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, body=clear_values_request_body)
    response = request.execute()
    
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED",
                                    body={"values": LIST_SOCIEDADES_POR_CIDADE_NAOPAGO}).execute()  
#%%

STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro, 
a1.[Codigo Debito] as Debito,
a2.descricao, 
a1.Parcela as Parcela,
a1.[Data Base] as [Data Base],
a1.[Data Vencimento] as [Data Vencimento], 
a1.[Dt Notificacao] as [Data Notificacao],
a1.[Data Execucao Judicial] as [Data Execucao],
a1.[Dt Protesto] as [Data Protesto], 
a1.[Valor Originario] as [Valor Originario], 
a1.[Valor C.Monetaria] as [Correcao Monetaria], 
a1.[Valor Multa] as Multa, 
a1.[Valor Juros] as Juros, 
a1.[Valor Total] as Total 
from VIEW_SFN_SFNA01_CORRIGIDO a1, SFNT108 a2 
where 
a1.[Codigo Debito] = a2.[Codigo Debito]
order by a1.[Num. Registro],a1.[Codigo Debito]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
ESTOQUE_PROF = pd.DataFrame(query)

DEVEDORES=ESTOQUE_PROF

ESTOQUE_PROF_PARC = pd.DataFrame()
ESTOQUE_PROF_PARC=ESTOQUE_PROF.loc[ESTOQUE_PROF['Parcela']!=0]
ESTOQUE_PROF=ESTOQUE_PROF.drop(ESTOQUE_PROF.loc[ESTOQUE_PROF['Parcela']!=0].index, inplace=False)
ESTOQUE_PROF=ESTOQUE_PROF.reset_index(drop=True)


ESTOQUE_PROF= ESTOQUE_PROF.astype(str)
ESTOQUE_PROF['Valor Originario'] = ESTOQUE_PROF['Valor Originario'].str.replace('.',',', regex=True)
ESTOQUE_PROF['Correcao Monetaria'] = ESTOQUE_PROF['Correcao Monetaria'].str.replace('.',',', regex=True)
ESTOQUE_PROF['Multa'] = ESTOQUE_PROF['Multa'].str.replace('.',',', regex=True)
ESTOQUE_PROF['Juros'] = ESTOQUE_PROF['Juros'].str.replace('.',',', regex=True)
ESTOQUE_PROF['Total'] = ESTOQUE_PROF['Total'].str.replace('.',',', regex=True)
LIST_ESTOQUE_PROF= [ESTOQUE_PROF.columns.values.tolist()] + ESTOQUE_PROF.values.tolist()
ESTOQUE_PROF_sheets() 

ESTOQUE_PROF_PARC= ESTOQUE_PROF_PARC.astype(str)
ESTOQUE_PROF_PARC['Valor Originario'] = ESTOQUE_PROF_PARC['Valor Originario'].str.replace('.',',', regex=True)
ESTOQUE_PROF_PARC['Correcao Monetaria'] = ESTOQUE_PROF_PARC['Correcao Monetaria'].str.replace('.',',', regex=True)
ESTOQUE_PROF_PARC['Multa'] = ESTOQUE_PROF_PARC['Multa'].str.replace('.',',', regex=True)
ESTOQUE_PROF_PARC['Juros'] = ESTOQUE_PROF_PARC['Juros'].str.replace('.',',', regex=True)
ESTOQUE_PROF_PARC['Total'] = ESTOQUE_PROF_PARC['Total'].str.replace('.',',', regex=True)

LIST_ESTOQUE_PROF_PARC= [ESTOQUE_PROF_PARC.columns.values.tolist()] + ESTOQUE_PROF_PARC.values.tolist()
# ESTOQUE_PROF_PARC_sheets()



STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro, 
a1.[Codigo Debito] as Debito,
a2.descricao, 
a1.Parcela as Parcela,
a1.[Data Base] as [Data Base],
a1.[Data Vencimento] as [Data Vencimento], 
a1.[Dt Notificacao] as [Data Notificacao],
a1.[Data Execucao Judicial] as [Data Execucao], 
a1.[Dt Protesto] as [Data Protesto], 
a1.[Valor Originario] as [Valor Originario], 
a1.[Valor C.Monetaria] as [Correcao Monetaria], 
a1.[Valor Multa] as Multa, 
a1.[Valor Juros] as Juros, 
a1.[Valor Total] as Total 
from VIEW_SFN_SFNA02_CORRIGIDO a1, SFNT108 a2 
where 
a1.[Codigo Debito] = a2.[Codigo Debito] 
order by a1.[Num. Registro],a1.[Codigo Debito]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
ESTOQUE_SOC = pd.DataFrame(query)
DEVEDORES_SOC=ESTOQUE_SOC
PERFIL_SOC=ESTOQUE_SOC.loc[ESTOQUE_SOC['Parcela']==0]
ESTOQUE_SOC = ESTOQUE_SOC.astype(str)

ESTOQUE_SOC['Valor Originario'] = ESTOQUE_SOC['Valor Originario'].str.replace('.',',', regex=True)
ESTOQUE_SOC['Correcao Monetaria'] = ESTOQUE_SOC['Correcao Monetaria'].str.replace('.',',', regex=True)
ESTOQUE_SOC['Multa'] = ESTOQUE_SOC['Multa'].str.replace('.',',', regex=True)
ESTOQUE_SOC['Juros'] = ESTOQUE_SOC['Juros'].str.replace('.',',', regex=True)
ESTOQUE_SOC['Total'] = ESTOQUE_SOC['Total'].str.replace('.',',', regex=True)
LIST_ESTOQUE_SOC= [ESTOQUE_SOC.columns.values.tolist()] + ESTOQUE_SOC.values.tolist()
ESTOQUE_SOC_sheets() 


STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro, 
a1.[Codigo Debito] as Debito,
a2.descricao, 
a1.Parcela as Parcela,
a1.[Data Base] as [Data Base],
a1.[Data Vencimento] as [Data Vencimento], 
a1.[Dt Notificacao] as [Data Notificacao],
a1.[Data Execucao Judicial] as [Data Execucao], 
a1.[Dt Protesto] as [Data Protesto], 
a1.[Valor Originario] as [Valor Originario], 
a1.[Valor C.Monetaria] as [Correcao Monetaria], 
a1.[Valor Multa] as Multa, 
a1.[Valor Juros] as Juros, 
a1.[Valor Total] as Total 
from VIEW_SFN_SFNA04_CORRIGIDO a1, SFNT108 a2 
where 
a1.[Codigo Debito] = a2.[Codigo Debito] 
order by a1.[Num. Registro],a1.[Codigo Debito]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
ESTOQUE_PFK = pd.DataFrame(query)
ESTOQUE_PFK = ESTOQUE_PFK.astype(str)

ESTOQUE_PFK['Valor Originario'] = ESTOQUE_PFK['Valor Originario'].str.replace('.',',', regex=True)
ESTOQUE_PFK['Correcao Monetaria'] = ESTOQUE_PFK['Correcao Monetaria'].str.replace('.',',', regex=True)
ESTOQUE_PFK['Multa'] = ESTOQUE_PFK['Multa'].str.replace('.',',', regex=True)
ESTOQUE_PFK['Juros'] = ESTOQUE_PFK['Juros'].str.replace('.',',', regex=True)
ESTOQUE_PFK['Total'] = ESTOQUE_PFK['Total'].str.replace('.',',', regex=True)
LIST_ESTOQUE_PFK= [ESTOQUE_PFK.columns.values.tolist()] + ESTOQUE_PFK.values.tolist()
ESTOQUE_PFK_sheets() 


STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro, 
a1.[Codigo Debito] as Debito,
a2.descricao, 
a1.Parcela as Parcela,
a1.[Data Base] as [Data Base],
a1.[Data Vencimento] as [Data Vencimento], 
a1.[Dt Notificacao] as [Data Notificacao],
a1.[Data Execucao Judicial] as [Data Execucao],
a1.[Dt Protesto] as [Data Protesto],  
a1.[Valor Originario] as [Valor Originario], 
a1.[Valor C.Monetaria] as [Correcao Monetaria], 
a1.[Valor Multa] as Multa, 
a1.[Valor Juros] as Juros, 
a1.[Valor Total] as Total 
from VIEW_SFN_SFNA05_CORRIGIDO a1, SFNT108 a2 
where 
a1.[Codigo Debito] = a2.[Codigo Debito] 
order by a1.[Num. Registro],a1.[Codigo Debito]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
ESTOQUE_PJK = pd.DataFrame(query)
ESTOQUE_PJK = ESTOQUE_PJK.astype(str)

ESTOQUE_PJK['Valor Originario'] = ESTOQUE_PJK['Valor Originario'].str.replace('.',',', regex=True)
ESTOQUE_PJK['Correcao Monetaria'] = ESTOQUE_PJK['Correcao Monetaria'].str.replace('.',',', regex=True)
ESTOQUE_PJK['Multa'] = ESTOQUE_PJK['Multa'].str.replace('.',',', regex=True)
ESTOQUE_PJK['Juros'] = ESTOQUE_PJK['Juros'].str.replace('.',',', regex=True)
ESTOQUE_PJK['Total'] = ESTOQUE_PJK['Total'].str.replace('.',',', regex=True)
LIST_ESTOQUE_PJK= [ESTOQUE_PJK.columns.values.tolist()] + ESTOQUE_PJK.values.tolist()
ESTOQUE_PJK_sheets() 

##JUNÇÃO DOS FRAMES
# TRANST = [ESTOQUE_PROF, ESTOQUE_SOC,ESTOQUE_PFK,ESTOQUE_PJK]
# TOTAL=pd.concat(TRANST)
# TOTAL.to_excel('TOTAL.xlsx', index = False)


STRCOMANDO=str("""set dateformat dmy select distinct
               a1.[Numero Guia],a1.[Data Pagamento], a1.[Vlr Pago Total]
               from SFNH03 a1
               where 
               a1.[Data Pagamento] >= '01/01/2020' and
               a1.[Numero Guia] <>''""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
SFNH03 = pd.DataFrame(query)

STRCOMANDO=str("""set dateformat dmy select distinct
               a1.[Numero Guia],a1.[Data Pagamento], a1.[Valor Pago] as [Vlr Pago Total]
               from SFNH04 a1
               where 
               a1.[Data Pagamento] >= '01/01/2020' and
               a1.[Numero Guia] <>''""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
SFNH04 = pd.DataFrame(query)


STRCOMANDO=str("""set dateformat dmy select distinct
               a1.[Numero Guia],a1.[Data Lote] as [Data Pagamento], a1.[VALOR DO AVISO] as [Vlr Pago Total]
               from SFNH05 a1
               where 
               a1.[Data Lote] >= '01/01/2020' and
               a1.[Numero Guia] <>'' and
               a1.[VALOR DO AVISO]>'1'""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
SFNH05 = pd.DataFrame(query)


TRANST2 = [SFNH03,SFNH04,SFNH05]

PAGAMENTOS=pd.concat(TRANST2)
PAGAMENTOS=PAGAMENTOS.drop([ 'Numero Guia'], axis=1)
PAGAMENTOS=PAGAMENTOS.groupby("Data Pagamento",as_index=False)['Vlr Pago Total'].sum()
PAGAMENTOS=PAGAMENTOS.astype(str)
PAGAMENTOS['Vlr Pago Total'] = PAGAMENTOS['Vlr Pago Total'].str.replace('.',',', regex=True)
LIST_PAGAMENTOS= [PAGAMENTOS.columns.values.tolist()] + PAGAMENTOS.values.tolist()

PAGAMENTOS_sheets()



STRCOMANDO=str("""set dateformat dmy 
select distinct a1.[Num. Registro],a1.[Data Pagamento], a1.[Vlr Pago Total], a1.[Parcela]
from SFNH03 a1
where
a1.[Data Lote] >= '27/10/2022' and
a1.Parcela between '0' and '1' and
a1.[Num. Registro] in 
(select distinct a1.[Num. Registro] from SFNW11 a1
where
a1.[numero guia] between '28058500000717437' and '28058500000737887')""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
NOTIFICACAO = pd.DataFrame(query)
NOTIFICACAO= NOTIFICACAO.astype(str)
NOTIFICACAO['Vlr Pago Total'] = NOTIFICACAO['Vlr Pago Total'].str.replace('.',',', regex=True)
LIST_RESULT_NOTIFICACAO= [NOTIFICACAO.columns.values.tolist()] + NOTIFICACAO.values.tolist()
RESULT_NOTIFICAO_sheets() 



PERFIL=ESTOQUE_PROF
PERFIL=PERFIL.drop([ 'Debito', 'descricao', 'Parcela', 'Data Vencimento', 'Data Notificacao', 'Data Execucao','Data Protesto', 'Valor Originario', 'Correcao Monetaria', 'Multa','Juros'], axis=1)
PERFIL["Total"]=PERFIL["Total"].str.replace(',','.', regex=True)
PERFIL['Total']=PERFIL['Total'].astype(float)
PERFIL=PERFIL.groupby("Registro",as_index=False)['Total'].sum()
PERFIL['Total']=PERFIL['Total'].astype(str)
PERFIL["Total"]=PERFIL["Total"].str.replace('.',',', regex=True)
PERFIL=PERFIL.assign(Sexo="")
PERFIL=PERFIL.assign(DataNascimento="0")
PERFIL=PERFIL.assign(Categoria="0")
PERFIL=PERFIL.assign(Cidade="0")
PERFIL=PERFIL.assign(Bairro="0")
PERFIL=PERFIL.assign(Situacao="0")


linha=0
for n in PERFIL.index:

    registro=PERFIL['Registro'][linha]
    
    STRCOMANDO = str(
        "set dateformat dmy SELECT [Dt Nascimento], Sexo, Categoria,[Situacao Cadastral] FROM SCDA01 a1 where [Num. Registro]='"+registro+"'")
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server +
                          ';DATABASE='+database+';UID='+username+';PWD=' + password)
    cursor = cnxn.cursor()
    comando = STRCOMANDO
    query = pd.read_sql(comando, cnxn)
    cursor.execute(comando)
    cadastral = pd.DataFrame(query)
      
        
        
    if cadastral.empty:
        PERFIL['DataNascimento'][linha]='DESCONHECIDO'
        PERFIL['Sexo'][linha]='DESCONHECIDO'
        PERFIL['Categoria'][linha]='DESCONHECIDO'
        PERFIL['Situacao'][linha]='DESCONHECIDO'
    else:
        if cadastral["Situacao Cadastral"][0]==1:
            situacao='Ativo'
        elif cadastral["Situacao Cadastral"][0]==2 or cadastral["Situacao Cadastral"][0]==3 or cadastral["Situacao Cadastral"][0]==5 or cadastral["Situacao Cadastral"][0]==99:
            situacao='Baixado'
        elif cadastral["Situacao Cadastral"][0]==6 :
            situacao='Suspenso'    
        elif cadastral["Situacao Cadastral"][0]==7 or cadastral["Situacao Cadastral"][0]==8 or cadastral["Situacao Cadastral"][0]==9:
            situacao='Baixado'  
        PERFIL['DataNascimento'][linha]=cadastral["Dt Nascimento"][0]
        PERFIL['Sexo'][linha]=cadastral["Sexo"][0]
        PERFIL['Categoria'][linha]=cadastral["Categoria"][0]
        PERFIL['Situacao'][linha]=situacao
        
    STRCOMANDO = str(
        "set dateformat dmy SELECT Cidade, Bairro FROM SCDA51 a1 where [Num. Registro]='"+registro+"' and [Endereco Ativo]='SIM' and [Endereco Correspondencia]='SIM'")
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server +
                          ';DATABASE='+database+';UID='+username+';PWD=' + password)
    cursor = cnxn.cursor()
    comando = STRCOMANDO
    query = pd.read_sql(comando, cnxn)
    cursor.execute(comando)
    endereco = pd.DataFrame(query)
    if endereco.empty:
        PERFIL['Cidade'][linha]='DESCONHECIDO'
        PERFIL['Bairro'][linha]='DESCONHECIDO'
    else:
        PERFIL['Cidade'][linha]=endereco["Cidade"][0]
        PERFIL['Bairro'][linha]=endereco["Bairro"][0]
    linha+=1


PERFIL=PERFIL.astype(str)
PERFIL["Sexo"]=PERFIL["Sexo"].str.replace('2','FEMININO', regex=True)
PERFIL["Sexo"]=PERFIL["Sexo"].str.replace('1','MASCULINO', regex=True)
PERFIL["Categoria"]=PERFIL["Categoria"].str.replace('2','TECNICO', regex=True)
PERFIL["Categoria"]=PERFIL["Categoria"].str.replace('1','CONTADOR', regex=True)
PERFIL["Total"]=PERFIL["Total"].str.replace('.',',', regex=True)
#PERFIL["Cidade"]=PERFIL["Cidade"].str.replace('SAO GONCALO','SÃO GONÇALO', regex=True)
LIST_PERFIL= [PERFIL.columns.values.tolist()] + PERFIL.values.tolist()
PERFIL_sheets() 

PERFIL_SOC=PERFIL_SOC.astype(str)
PERFIL_SOC=PERFIL_SOC.drop([ 'Debito', 'descricao', 'Parcela', 'Data Vencimento', 'Data Notificacao', 'Data Execucao','Data Protesto', 'Valor Originario', 'Correcao Monetaria', 'Multa','Juros'], axis=1)
PERFIL_SOC["Total"]=PERFIL_SOC["Total"].str.replace(',','.', regex=True)
PERFIL_SOC['Total']=PERFIL_SOC['Total'].astype(float)
PERFIL_SOC=PERFIL_SOC.groupby("Registro",as_index=False)['Total'].sum()
PERFIL_SOC['Total']=PERFIL_SOC['Total'].astype(str)
PERFIL_SOC["Total"]=PERFIL_SOC["Total"].str.replace('.',',', regex=True)


PERFIL_SOC=PERFIL_SOC.assign(SituacaoCadastral="")
PERFIL_SOC=PERFIL_SOC.assign(TipoSociedade="")    
PERFIL_SOC=PERFIL_SOC.assign(EmpregadoQuantidade="")        
PERFIL_SOC=PERFIL_SOC.assign(DtRegistroOriginario="")      
PERFIL_SOC=PERFIL_SOC.assign(Cidade="")
PERFIL_SOC=PERFIL_SOC.assign(Bairro="")


linha=0
for n in PERFIL_SOC.index:

    registro=PERFIL_SOC['Registro'][linha]
    
    STRCOMANDO = str(
    "set dateformat dmy SELECT [Situacao Cadastral],[Tipo de Sociedade], [Empregado Quantidade], [Dt Registro Originario] FROM SCDA02 a1 where [Num. Registro]='"+registro+"'")
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server +
                          ';DATABASE='+database+';UID='+username+';PWD=' + password)
    cursor = cnxn.cursor()
    comando = STRCOMANDO
    query = pd.read_sql(comando, cnxn)
    cursor.execute(comando)
    cadastral = pd.DataFrame(query)
    
    if cadastral.empty:
        PERFIL_SOC['SituacaoCadastral'][linha]="DESCONHECIDO"

        PERFIL_SOC['TipoSociedade'][linha]="DESCONHECIDO"
        
        PERFIL_SOC['EmpregadoQuantidade'][linha]="DESCONHECIDO"
        
        PERFIL_SOC['DtRegistroOriginario'][linha]="DESCONHECIDO"
        
      

    else:
        PERFIL_SOC['SituacaoCadastral'][linha]=cadastral['Situacao Cadastral'][0]

        PERFIL_SOC['TipoSociedade'][linha]=cadastral['Tipo de Sociedade'][0]
        
        PERFIL_SOC['EmpregadoQuantidade'][linha]=cadastral['Empregado Quantidade'][0]
        
        PERFIL_SOC['DtRegistroOriginario'][linha]=cadastral['Dt Registro Originario'][0]
        
        
    STRCOMANDO = str(
        "set dateformat dmy SELECT Cidade, Bairro FROM SCDA52 a1 where [Num. Registro]='"+registro+"'")
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server +
                          ';DATABASE='+database+';UID='+username+';PWD=' + password)
    cursor = cnxn.cursor()
    comando = STRCOMANDO
    query = pd.read_sql(comando, cnxn)
    cursor.execute(comando)
    endereco = pd.DataFrame(query)
    if endereco.empty:
        PERFIL_SOC['Cidade'][linha]='DESCONHECIDO'
        PERFIL_SOC['Bairro'][linha]='DESCONHECIDO'
    else:
        PERFIL_SOC['Cidade'][linha]=endereco["Cidade"][0]
        PERFIL_SOC['Bairro'][linha]=endereco["Bairro"][0]
    linha+=1

PERFIL_SOC=PERFIL_SOC.astype(str)
PERFIL_SOC["SituacaoCadastral"]=PERFIL_SOC["SituacaoCadastral"].str.replace('1','ATIVO', regex=True)
baixados = '|'.join(['2', '3', '5','99'])
PERFIL_SOC["SituacaoCadastral"]=PERFIL_SOC["SituacaoCadastral"].str.replace(baixados,'BAIXADOS', regex=True)
PERFIL_SOC["SituacaoCadastral"]=PERFIL_SOC["SituacaoCadastral"].str.replace('6','SUSPENSO', regex=True)
cancelados = '|'.join(['7', '8', '9','31','33'])
PERFIL_SOC["SituacaoCadastral"]=PERFIL_SOC["SituacaoCadastral"].str.replace(cancelados,'CANCELADOS', regex=True)


PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1000','INDIVIDUAL.', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1001','SOCIEDADE EMPRESÁRIA LTDA', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1002','SOCIEDADE SIMPLES PURA', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1003','SOCIEDADE SIMPLES LTDA', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1004','EMPRESÁRIO(INDIVIDUAL)', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1005','MEI', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1006','EIRELI', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1007','SOCIEDADE SIMPLES EM NOME COLETIVO', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1008','SOCIEDADE LIMITADA UNIPESSOAL (SLU)', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('1','SOC.PROF.', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('2','SOC.MISTA', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('3','INDIVIDUAL', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('4','CEI EQUIP. A PJ', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('5','AUDITORIA INDEPENDENTE', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('6','AUDITORIA E CONTABILIDADE', regex=True)
PERFIL_SOC["TipoSociedade"]=PERFIL_SOC["TipoSociedade"].str.replace('7','PRAZO 180 DIAS ART.1033 CC', regex=True)
PERFIL_SOC["Cidade"]=PERFIL_SOC["Cidade"].str.replace('SAO GONCALO','SÃO GONÇALO', regex=True)


LIST_PERFIL_SOC= [PERFIL_SOC.columns.values.tolist()] + PERFIL_SOC.values.tolist()
PERFIL_SOC_sheets() 



STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro 

from SCDA01 a1 
where 
a1.[Tipo Situacao] <> 'S' and
a1.[Situacao Cadastral]='1'
order by a1.[Num. Registro]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
ATIVOS = pd.DataFrame(query)


STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro, 
a1.[Codigo Debito] as Debito,
a1.[Data Vencimento] as [Data Vencimento]

from VIEW_SFN_SFNA01_CORRIGIDO a1, SFNT108 a2 
where 
a1.[Codigo Debito] = a2.[Codigo Debito] and
a1.[Codigo Debito] like '242%' AND
a1.[Parcela]='0'
order by a1.[Num. Registro],a1.[Codigo Debito]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
DEVEDORES = pd.DataFrame(query)



ATIVO_DEV = DEVEDORES.loc[(DEVEDORES['Registro'].isin(ATIVOS['Registro']))].reset_index(drop=True)
ATIVO_DEV['Data Vencimento']=pd.to_datetime(ATIVO_DEV['Data Vencimento'],format='%Y-%m-%d', errors='coerce')
ATIVO_DEV['Data Vencimento'] = ATIVO_DEV['Data Vencimento'].dt.date
ATIVO_DEV=ATIVO_DEV.loc[(ATIVO_DEV['Data Vencimento']<=HOJE)].reset_index(drop=True)

TESTE=pd.DataFrame({"TOTAL_ATIVOS_DEVEDORES":[len(ATIVO_DEV['Registro'])],
                    "TOTAL_ATIVOS":[len(ATIVOS['Registro'])]})

LIST_TOTAL_ATIVOS= [TESTE.columns.values.tolist()] + TESTE.values.tolist()
TOTAL_ATIVOS_PROF_sheets()


STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro 

from SCDA02 a1 
where 
a1.[Tipo Situacao] <> 'S' and
a1.[Situacao Cadastral]='1'
order by a1.[Num. Registro]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
ATIVOS_SOC = pd.DataFrame(query)


STRCOMANDO=str("""select 
a1.[Num. Registro] as Registro, 
a1.[Codigo Debito] as Debito,
a1.[Data Vencimento] as [Data Vencimento]

from VIEW_SFN_SFNA02_CORRIGIDO a1, SFNT108 a2 
where 
a1.[Codigo Debito] = a2.[Codigo Debito] and
a1.[Codigo Debito] like '243%' AND
a1.[Parcela]='0'
order by a1.[Num. Registro],a1.[Codigo Debito]""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
DEVEDORES_SOC = pd.DataFrame(query)

DEVEDORES_SOC=DEVEDORES_SOC.groupby("Registro",as_index=False)['Data Vencimento'].min()

ATIVO_DEV_SOC=DEVEDORES_SOC.loc[(DEVEDORES_SOC['Registro'].isin(ATIVOS_SOC['Registro']))].reset_index(drop=True)
ATIVO_DEV_SOC['Data Vencimento']=pd.to_datetime(ATIVO_DEV_SOC['Data Vencimento'],format='%Y-%m-%d', errors='coerce')
ATIVO_DEV_SOC['Data Vencimento'] = ATIVO_DEV_SOC['Data Vencimento'].dt.date
ATIVO_DEV_SOC=ATIVO_DEV_SOC.loc[(ATIVO_DEV_SOC['Data Vencimento']<=HOJE)].reset_index(drop=True)

TESTE2=pd.DataFrame({"TOTAL_ATIVOS_DEVEDORES":[len(ATIVO_DEV_SOC['Registro'])],
                    "TOTAL_ATIVOS":[len(ATIVOS_SOC['Registro'])]})

LIST_TOTAL_SOC_ATIVOS= [TESTE2.columns.values.tolist()] + TESTE2.values.tolist()
TOTAL_ATIVOS_SOC_sheets()


STRCOMANDO=str("""set dateformat dmy select distinct
                a1.[Num. Registro],a1.[Cidade], a1.[Total Guia]
               from SFNW11 a1, SFNW12 a2
               where
               a1.[Numero Guia]=a2.[Numero Guia] and
               a1.Retorno='NAO' and
               a1.[Data Impressao] > '06/12/2022' and
               a1.[Numero Guia] <>'' and
               a1.[Tipo Registro]='2' and
               a2.[Parcela] BETWEEN '0' AND '1' and
               a2. [Tipo Debito] ='1'""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
SOCIEDADES_POR_CIDADE_NAOPAGO = pd.DataFrame(query)

STRCOMANDO=str("""set dateformat dmy select distinct
                a1.[Num. Registro],a1.[Cidade], a1.[Total Guia]
               from SFNW11 a1, SFNW12 a2
               where
               a1.[Numero Guia]=a2.[Numero Guia] and
               a1.Retorno='SIM' and
               a1.[Data Impressao] > '06/12/2022' and
               a1.[Numero Guia] <>'' and
               a1.[Tipo Registro]='2' and 
               a2.[Parcela] BETWEEN '0' AND '1' and
               a2. [Tipo Debito] ='1'""")

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
comando=STRCOMANDO
query = pd.read_sql(comando, cnxn)
cursor.execute(comando)
SOCIEDADES_POR_CIDADE_PAGO = pd.DataFrame(query)

SOCIEDADES_POR_CIDADE_PAGO=SOCIEDADES_POR_CIDADE_PAGO.astype(str)
SOCIEDADES_POR_CIDADE_NAOPAGO=SOCIEDADES_POR_CIDADE_NAOPAGO.astype(str)
SOCIEDADES_POR_CIDADE_PAGO["Total Guia"]=SOCIEDADES_POR_CIDADE_PAGO["Total Guia"].str.replace('.',',', regex=True)
SOCIEDADES_POR_CIDADE_NAOPAGO["Total Guia"]=SOCIEDADES_POR_CIDADE_NAOPAGO["Total Guia"].str.replace('.',',', regex=True)

LIST_SOCIEDADES_POR_CIDADE_PAGO=[SOCIEDADES_POR_CIDADE_PAGO.columns.values.tolist()] + SOCIEDADES_POR_CIDADE_PAGO.values.tolist()
SOCIEDADES_POR_CIDADE_PAGO_sheets()

LIST_SOCIEDADES_POR_CIDADE_NAOPAGO= [SOCIEDADES_POR_CIDADE_NAOPAGO.columns.values.tolist()] + SOCIEDADES_POR_CIDADE_NAOPAGO.values.tolist()
# SOCIEDADES_POR_CIDADE_NAOPAGO_sheets()

cnxn.close()

# concluido()



