from time import sleep
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
from xml.dom import minidom
import os.path
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import *
from tkinter import messagebox


try:
    
    
        
    print("Executando o código que apaga as linhas da planilha 'Pedidos de compra'")    
    #Aqui irá rodar o código de validação do Google sheet, a validação de suas credenciais
    #Com essa função é possível rodar o script que abre os arquivos xml e buscar as informações dentro de cada  
    # nota fiscal.
    sleep(5)
    #Varíavel para a localização das planilhas Google
    scope = ['https://spreadsheets.google.com/feeds']


    #Dados de autenticação
    #Computador trabalho
    #credentials = ServiceAccountCredentials.from_json_keyfile_name('E:\Albacete-automation\Credential_google.json', scope)

    #Computador casa
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Albacete-automation/Credential_google.json', scope)


    #Se autentica
    gc = gspread.authorize(credentials)


    #Abre a planilha
    wks = gc.open_by_key('1Sy5HQBbSRewZrr3ZLtCaCHRqVul2ihajVoLscePMpVw')


    #Para selecionar a planilha pelo o nome use o código abaixo
    #wks = gc.open("Planejamento - Motores WEG") 

    #Seleciona a primeira página da planilha
    #worksheet = wks.get_worksheet(18)
    #worksheet = wks.get_worksheet(19)
    worksheet = wks.worksheet("Pedidos de Compra") 

    dados = worksheet.get_all_records()

    planilha01 = pd.read_excel("C:/Albacete-automation/DATABASE/Parametros-dos-motores.xlsx", sheet_name="Parâmetros dos Motores")

    invoiceArray = []
    searchPath = 'C:/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE'
    def delete_rows():
        for file in os.listdir("C:/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE"):
                        if file.endswith(".xml"):
                                #invoiceArray = [os.path.join(file)]
                                invoiceArray.append(os.path.join(file))
                                
                                

                                pass
                
        i= -1
        while True:
                        weg = True
                        
                        for invoice in invoiceArray:
                                i += 1    
                                if os.path.exists(searchPath+"/"+invoice):
                                        #print(i)
                                        with open(searchPath+"/"+invoice, 'r', encoding='utf-8') as f:
                                                xml = minidom.parse(f)
                                                nf = xml.getElementsByTagName("nNF")
                                                clienteOrder = xml.getElementsByTagName("xPed")
                                                time = xml.getElementsByTagName("dhEmi")
                                                productCode = xml.getElementsByTagName("cProd")
                                                amountOrdered = xml.getElementsByTagName("qCom")
                                                
                        #Utilizada para obter o número da nota fiscal do xml    
                                        for tag in nf:
                                                nota = [(tag.firstChild.data)]

                                #Utilizada para obter o número ordem de compra da ALbacete do xml          
                                        clienteOrderLoop = []
                                        for tag in clienteOrder:
                                #variável receve o valor encontrado no xml
                                                clienteOrder = (tag.firstChild.data)
                                                clienteOrderLoop.append(clienteOrder)
                                                print(clienteOrderLoop)
                                                print(len(clienteOrderLoop))
                                                
                                        for x in clienteOrderLoop:
                                                
                                                #aqui fica o "for" responsável por deletar da planilha 
                                                #Aqui printa todas as Ordens de Compra
                                                print(x)
                                                #Ela é encontrada e traz o objeto com os dados de localização
                                                #Exemplo: <Cell R53C1 '035141'>, onde Cell é a célula
                                                #R53 é row53 ou linha 53 e C1 é Column1 ou Coluna 1
                                                cell_list = worksheet.find(x)
                                                #cell_list = worksheet.find('030657')
                                                #converter classe para string para obter o tamanho
                                                conversion = str(cell_list)
                                                print(len(conversion))
                                                if len(conversion)==20:
                                                        print(cell_list)
                                                        # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                        #Exemplo: 53 or 55 
                                                        cell_list = str(cell_list)[7:8] 
                                                        #Converter para inteiro 
                                                        rowShow = int(cell_list)
                                                        #rowShow = cell_list - Linha alternativa
                                                        print("Esta é a linha:",rowShow)
                                                        #Deleta a linha encontrada
                                                        worksheet.delete_rows(rowShow)
                                                        pass
                                                if len(conversion)==21:
                                                        print(cell_list)
                                                        # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                        #Exemplo: 53 or 55 
                                                        cell_list = str(cell_list)[7:9] 
                                                        #Converter para inteiro 
                                                        rowShow = int(cell_list)
                                                        #rowShow = cell_list - Linha alternativa
                                                        print("Esta é a linha:",rowShow)
                                                        #Deleta a linha encontrada
                                                        worksheet.delete_rows(rowShow)
                                                        pass
                                                if len(conversion)==22:
                                                        print(cell_list)
                                                        # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                        #Exemplo: 53 or 55 
                                                        cell_list = str(cell_list)[7:10] 
                                                        #Converter para inteiro 
                                                        rowShow = int(cell_list)
                                                        #rowShow = cell_list - Linha alternativa
                                                        print("Esta é a linha:",rowShow)
                                                        #Deleta a linha encontrada
                                                        worksheet.delete_rows(rowShow)
                                                        pass
                                                if len(conversion)==23:
                                                        print(cell_list)
                                                        # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                        #Exemplo: 53 or 55 
                                                        cell_list = str(cell_list)[7:11] 
                                                        #Converter para inteiro 
                                                        rowShow = int(cell_list)
                                                        #rowShow = cell_list - Linha alternativa
                                                        print("Esta é a linha:",rowShow)
                                                        #Deleta a linha encontrada
                                                        worksheet.delete_rows(rowShow)
                                                        pass
                                                #depois de obtido o tamanho, iremos colocar uma condição para achar a linha exata de acordo com o tamanho
                                                print(cell_list)
                                                
                                                
                                                
                                                
                                                #Deleta a linha encontrada
                                                #worksheet.delete_rows(rowShow)
                                                
                                                
                                                if len(clienteOrderLoop) == 2:
                                                        print(clienteOrderLoop[1])
                                                        cell_list = worksheet.find(clienteOrderLoop[1])
                                                        conversionArray = str(cell_list)
                                                        if len(conversionArray)==20:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:8] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                                pass
                                                        if len(conversionArray)==21:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:9] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversionArray)==22:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:10] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversion)==23:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:11] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        
                                                if len(clienteOrderLoop) == 3:
                                                        print(clienteOrderLoop[2])
                                                        cell_list = worksheet.find(clienteOrderLoop[2])
                                                        conversionArray = str(cell_list)
                                                        if len(conversionArray)==20:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:8] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversionArray)==21:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:9] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversionArray)==22:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:10] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversion)==23:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:11] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        
                                                if len(clienteOrderLoop) ==4:
                                                        print(clienteOrderLoop[3])
                                                        cell_list = worksheet.find(clienteOrderLoop[3])
                                                        conversionArray = str(cell_list)
                                                        if len(conversionArray)==20:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:8] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversionArray)==21:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:9] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversionArray)==22:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:10] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversion)==23:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:11] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                
                                                if len(clienteOrderLoop) ==5:
                                                        print(clienteOrderLoop[4])
                                                        cell_list = worksheet.find(clienteOrderLoop[4])
                                                        conversionArray = str(cell_list)
                                                        if len(conversionArray)==20:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:8] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                                pass
                                                        if len(conversionArray)==21:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:9] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversionArray)==22:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:10] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        if len(conversion)==23:
                                                                print(cell_list)
                                                                # Abaixo conseguimos coletar a localização da linha onde ela se encontra
                                                                #Exemplo: 53 or 55 
                                                                cell_list = str(cell_list)[7:11] 
                                                                #Converter para inteiro 
                                                                rowShow = int(cell_list)
                                                                #rowShow = cell_list - Linha alternativa
                                                                print("Esta é a linha:",rowShow)
                                                                #Deleta a linha encontrada
                                                                worksheet.delete_rows(rowShow)
                                                        pass
                                                        
                                                pass
                                        
                                                #print("Found something at R%sC%s" % (cell_list1.row, cell_list1.col))
                                        
                                        if clienteOrder == []:
                                                clienteOrder = ['Not exist']

                                        elif clienteOrder == ['WMP - AMOSTRA']:
                                                        
                                                clienteOrder = ['Sample']
                                        else:
                                                pass
                                pass

                                
                        else:
                                print("Arquivo nao existe")
                                sleep(10)
                        weg = False
                        break

    delete_rows()       
                 
except FileNotFoundError as error:
      vmsg = "Não encontramos nem o arquivo e nem o diretório"
      tiposmg = error
      def showMessage(tiposmg, msg):
            if tiposmg == error:
                  messagebox.showerror(title="Sem arquivos ou diretório", message=msg)

      showMessage(tiposmg, vmsg )
      
except ValueError as error:
    vmsg = "Não encontramos mais nenhuma ordem de compra na planilha de 'Pedidos de Compra'"
    tiposmg = error
    def showMessage(tiposmg, msg):
            if tiposmg == error:
                  messagebox.showerror(title="Sem Pedido de Compra", message=msg)

    showMessage(tiposmg, vmsg )
    
#cell_list = worksheet.findall("Rug store")



    
"""
#importando um arquivo excel ou csv
df = pd.DataFrame(dados)
print("Tabela original: \n",df)

#Removendo duplicados
df = df.drop_duplicates()
df.dropna(how='all')
print('\n\n',df)
"""
"""
dados = worksheet.get_all_records()
df = pd.DataFrame(dados)
df.info()
print("Tabela original: \n",df)
df.dropna(how='all')
print(rows)
print(cols)
"""
"""
batata = worksheet.row_count

#worksheet.clear()
#Aqui conseguimos selecionar as cédulas para limpeza
#worksheet.batch_clear(["A2:F4"])
print([df.columns.values.tolist()] + df.values.tolist())
print("\n")
#print([df.columns.values.tolist()[15:]] + df.values.tolist())
print(df.values.tolist())
"""
"""
worksheet.delete_rows(1, batata-1)
print([df.columns.values.tolist()] + df.values.tolist())
worksheet.append_rows([df.columns.values.tolist()] + df.values.tolist())

"""


"""
df = df.sort_values('Age', ascending=False)
df = df.drop_duplicates(subset='Name', keep='first')
print(df)
"""