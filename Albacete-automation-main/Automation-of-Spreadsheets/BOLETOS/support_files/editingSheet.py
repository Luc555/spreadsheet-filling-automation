from selenium import webdriver
#Pandas é uma biblioteca para manipulação de dados de planilhas e importamos criando um 
#apelido que é o pd
import pandas as pd
from tkinter import *
#Biblioteca para manipulação de dados de planilhas do GoogleSheets
import gspread
import subprocess
from oauth2client.service_account import ServiceAccountCredentials

#Varíavel recebe uma lista vazia
ourCode = []
#Variável recebe um dataframe de abertura da planilha excel, aberta no diretório passado mais abaixo
planilha01 = pd.read_excel("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/planilha_weg/export.xls", sheet_name="open_duplicate")


#Testes de exibição do DataFrame
'''
#Exibe todo o dataframe
print(planilha01)
#Exibe a coluna 'Nosso Número'
print(planilha01[['Nosso Número']])
#Exibe as informações do dataframe
print(planilha01.info())
#Exibe as linhas com os valores vazios 
print(planilha01.isna())
'''

#Aqui ele remove as linhas que contem um valor vazio
#Aqui não específica a coluna
# AXIS siginifica a removação de linhas ou colunas, sendo 0 para linha e 1 para coluna
#df_remove = planilha01.dropna(how="any", axis=0)

#Aqui remove as linhas onde tem valores nulos na coluna 'Nosso Número'
# AXIS siginifica a removação de linhas ou colunas, sendo 0 para linha e 1 para coluna
df_remove = planilha01.dropna(subset=['Nosso Número'], axis = 0)

#Cria uma nova planilha com base no dataframe "df_remove" criado logo acima.
df = df_remove.to_excel('C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/planilha_weg/Atualizada.xlsx', index = False)

#Criação de um novo dataframe após a exclusão das colunas listadas abaixo, através do método
#dropna e AXIS aqui é igual a 1
df_remove_column = df_remove.drop(['Número da Duplicata', 'Setor de Atividade', 'Dias de Atraso', 'Data de Emissão', 'Nosso Número', 'CNPJ Empresa Emissora', 'CNPJ do Cliente', 'Número Cliente', 'Cliente', 'Moeda', 'Situação'], axis=1)

#Cria uma nova planilha com base no dataframe "df_remove_column" criado logo acima.
df = df_remove_column.to_excel('C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/planilha_weg/AtualizadaColunas.xlsx', index = False)

# Aqui somamos as colunas "Nota Fiscal"  + a coluna "Número Parcela" com o uso do '-' entre eles
df_remove_column['Nota Fiscal'] = df_remove_column['Nota Fiscal'].map(str) + '-' + df_remove_column['Número Parcela'].map(str)

#Transforma a coluna 'Nota Fiscal' do dataframe em uma lista
lista = df_remove_column['Nota Fiscal'].tolist()

#Cria a lista para armazenar as faturas(notas fiscais + número parcela)
faturaList = []
#Cria a lista para armazenar as notas fiscais
nfList = []

#Cria um laço para que a cada item na lista(Nota Fiscal)
#Variável 'item' recebe  o número da parcela + nota fiscal através da manipulação de caracteres
#Variável 'nf' recebe a  nota fiscal
#Transforma as strings em lista 
#Envia para as lista criadas mais acima.
for item in lista:
    item = item[0:-4]+item[len(item)-2:]
    nf = item[0:len(item)-2]
    item.split()
    nf.split()
    faturaList.append(item)
    nfList.append(nf)

#Atualiza a coluna "Número Parcela" com a lista 'faturaList'
df_remove_column['Número Parcela'] = faturaList
#Atualiza a coluna "Nota Fiscal" com a lista 'nfList'
df_remove_column['Nota Fiscal'] = nfList

#Reorganiza o dataframe para que as colunas sejam iguais
df_remove_column = df_remove_column[['Nota Fiscal', 'Número Parcela', 'Nome Empresa Emissora','Data de Vencimento','Valor da Duplicata' ]]
#Printando o dataframe final
print(df_remove_column)


#Varíavel para a localização das planilhas Google
scope = ['https://spreadsheets.google.com/feeds']



#Credencial de autenticação da planilha
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Albacete-automation/Credential_google.json', scope)

#Se autentica
gc = gspread.authorize(credentials)


#Abre a planilha através do link
wks = gc.open_by_key('1Sy5HQBbSRewZrr3ZLtCaCHRqVul2ihajVoLscePMpVw')

#Seleciona a primeira página da planilha
#worksheet = wks.get_worksheet(14)
#Seleciona a aba pelo nome
worksheet = wks.worksheet("Boletos") 

#worksheet.delete_rows(2)
#Linha que apaga as linhas da planilha a partir da terceira linha, pois nesta planilha
#além do menu, a última linha restante não pode ser apagada. 
#worksheet.row_count é a mesma coisa que o indice da planilha, ou seja, o total de linhas
worksheet.delete_rows(3, worksheet.row_count)

#Transforma os valores do dataframe 'df_remove_column' em uma lista, separando em lista completas
df = df_remove_column.values.tolist()
#Exibindo a lista recém criada
print(df)

#Interador inicia no zero
i=0
#Para cada linha na lista df(lembrando que é uma lista composta de listas)
for row in df:
    #Insere na planilha 'BOLETO' a linha com a lista extraída de dentro da lista 'df' a medida em que 
    # soma o interador
    worksheet.append_row(df[i])
    #Soma mais um à variável interador
    i=i+1

#Apaga a linha repetida, que não havia sido deletada anteriormente    
worksheet.delete_rows(2)



