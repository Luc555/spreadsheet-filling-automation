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
#wks = gc.open('Teste Python') 

#Seleciona a primeira página da planilha
#worksheet = wks.get_worksheet(14)
worksheet = wks.worksheet("Boletos") 

print(worksheet)
print(worksheet.row_count)


row = 1
while row<=worksheet.row_count+1:
    #print(worksheet.row_values(row))
    row = row+1
    print(row)
    #worksheet.delete_rows(2)
    worksheet.delete_rows(3, worksheet.row_count)


print(worksheet.row_count)

    

'''
list_of_hashes = worksheet.get_all_values()
print(list_of_hashes)
'''