import gspread
import subprocess
from oauth2client.service_account import ServiceAccountCredentials
 
 #Varíavel para a localização das planilhas Google
scope = ['https://spreadsheets.google.com/feeds']


#Dados de autenticação
#Computador trabalho
#credentials = ServiceAccountCredentials.from_json_keyfile_name('E:\Albacete-automation\Credential_google.json', scope)

#Computador casa
credentials = ServiceAccountCredentials.from_json_keyfile_name('E:/Albacete-Automation/Credential_google.json', scope)


#Se autentica
gc = gspread.authorize(credentials)


#Abre a planilha
wks = gc.open_by_key('1Sy5HQBbSRewZrr3ZLtCaCHRqVul2ihajVoLscePMpVw')

#Para selecionar a planilha pelo o nome use o código abaixo

batata = wks.worksheet("Entregas Solicitadas") 
print('BATATA',batata)
#Seleciona a primeira página da planilha
worksheet = wks.get_worksheet(15)
print('WORKSHEET',worksheet)
