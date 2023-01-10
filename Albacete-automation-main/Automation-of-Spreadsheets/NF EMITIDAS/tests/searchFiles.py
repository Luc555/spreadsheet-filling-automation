import os
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os.path
import numpy as np
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials


#Aqui irá rodar o código de validação do Google sheet, a validação de suas credenciais
# nota fiscal.
        
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
worksheet = wks.get_worksheet(15)

#Código que lê a planilha referenciada e sua respectiva aba
#Computador casa
planilha01 = pd.read_excel("E:/Albacete-automation/DATABASE/Parametros-dos-motores.xlsx", sheet_name="Parâmetros dos Motores")

wegInvoice = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice.xml'
wegInvoice0 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice0.xml'
wegInvoice1 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice1.xml'
wegInvoice2 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice2.xml'
wegInvoice3 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice3.xml'
wegInvoice4 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice4.xml'
wegInvoice5 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice5.xml'
wegInvoice6 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice6.xml'
wegInvoice7 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice7.xml'
wegInvoice8 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice8.xml'
wegInvoice9 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice9.xml'
wegInvoice10 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice10.xml'
wegInvoice11 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice11.xml'
wegInvoice12 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice12.xml'
wegInvoice13 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice13.xml'
wegInvoice14 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice14.xml'
wegInvoice15 = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice/wegInvoice15.xml'




searchPath = 'E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice'

for file in os.listdir("E:/ALBACETE-AUTOMATION/Automation-of-Spreadsheets/weg-invoice"):
    if file.endswith(".xml"):
#       print(os.path.join(file))
#        list = [os.path.join(file)]
        pass



#Aqui ele pede uma interação do usuário para continuar o código.
print('\n')                       
str(input("Digite Enter Para continuar"))
print('\n')                       

while True:
        weg = True
        
        if os.path.exists(wegInvoice0):
                with open(wegInvoice0, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
#variável receve o valor encontrado no xml
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
                        delivery = Datelist2[0]

# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]
#Transforma a lista referência(string) em uma lista de inteiros 
                        valores = [int(val) for val in productCode]
                        ref = valores[0] # Primeiro valor da lista
                        procv = [ planilha01.loc[planilha01['Ref'] == int(ref), 'Código'].iloc[0], planilha01.loc[planilha01['Ref'] == int(ref), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref'] == int(ref), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref'] == int(ref), 'Transportadora'].iloc[0]]
                        code = [procv[0]]
                        product_name = [procv[1]]
                        delivery_day = [procv[2]]
                        sum = (int(procv[2]))
                        delivery = int(delivery)
                        delivery = delivery + sum
                        delivery = str(delivery)
                        delivery = [delivery]
                        delivery = delivery + Datelist1 + Datelist
                        final_delivery = ["/".join(delivery)]
                        shipping_company = [procv[3]]
               
#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0:2])]
                        
                invoiceZeroList = np.array([nota+clienteOrder+date+code+productCode+product_name+amountOrdered+delivery_day+final_delivery+shipping_company])
                invoiceZeroList=invoiceZeroList.flatten().tolist()
                
                worksheet.append_row(invoiceZeroList, value_input_option='USER_ENTERED')

                print(invoiceZeroList)
                pass
        
        if os.path.exists(wegInvoice1):
                with open(wegInvoice1, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
                        delivery = Datelist2[0]
                        
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
                        productCode = [(tag.firstChild.data[10:18])]
#Transforma a lista referência(string) em uma lista de inteiros 
                        valores = [int(val) for val in productCode]
                        ref = valores[0] # Primeiro valor da lista
                        procv = [ planilha01.loc[planilha01['Ref'] == int(ref), 'Código'].iloc[0], planilha01.loc[planilha01['Ref'] == int(ref), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref'] == int(ref), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref'] == int(ref), 'Transportadora'].iloc[0]]
                        code = [procv[0]]
                        product_name = [procv[1]]
                        delivery_day = [procv[2]]
                        sum = (int(procv[2]))
                        delivery = int(delivery)
                        delivery = delivery + sum
                        delivery = str(delivery)
                        delivery = [delivery]
                        delivery = delivery + Datelist1 + Datelist
                        final_delivery = ["/".join(delivery)]
                        shipping_company = [procv[3]]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0:2])]
                 
       
                invoiceOneList = np.array([nota+clienteOrder+date+code+productCode+product_name+amountOrdered+delivery_day+final_delivery+shipping_company])
                invoiceOneList=invoiceOneList.flatten().tolist() 
                


#Atualiza celula
                worksheet.append_row(invoiceOneList, value_input_option='USER_ENTERED')

                print(invoiceOneList)      
                pass

        if os.path.exists(wegInvoice2):
                with open(wegInvoice2, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando/
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceTwoList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceTwoList=invoiceTwoList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice3):
                with open(wegInvoice3, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                               
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                
                invoiceThreeList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceThreeList=invoiceThreeList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice4):
                with open(wegInvoice4, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /

                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceFourList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceFourList=invoiceFourList.flatten().tolist()
                pass   
        
        if os.path.exists(wegInvoice5):
                with open(wegInvoice5, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                        
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                
                invoiceFiveList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceFiveList=invoiceFiveList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice6):
                with open(wegInvoice6, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                               
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceSixList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceSixList=invoiceSixList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice7):
                with open(wegInvoice7, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceSevenList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceSevenList=invoiceSevenList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice8):
                with open(wegInvoice8, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceEightList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceEightList=invoiceEightList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice9):
                with open(wegInvoice9, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceNineList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceNineList=invoiceNineList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice10):
                with open(wegInvoice10, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceTenList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceTenList=invoiceTenList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice11):
                with open(wegInvoice11, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                invoiceElevenList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceElevenList=invoiceElevenList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice12):
                with open(wegInvoice12, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                
                invoiceTwelveList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceTwelveList = invoiceTwelveList.flatten().tolist()
                pass   

        if os.path.exists(wegInvoice13):
                with open(wegInvoice13, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceThirteenList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceThirteenList = invoiceThirteenList.flatten().tolist()
                pass   
                print('\n')

        if os.path.exists(wegInvoice14):
                with open(wegInvoice14, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceFourteenList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceFourteenList = invoiceFourteenList.flatten().tolist()
                pass   
                print('\n')

        if os.path.exists(wegInvoice15):
                with open(wegInvoice15, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]

#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceFifteenList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                invoiceFifteenList = invoiceFifteenList.flatten().tolist()
                pass   
                print('\n')
       
        if os.path.exists(wegInvoice):
                with open(wegInvoice, 'r', encoding='utf-8') as f:
                        xml = minidom.parse(f)
                        nf = xml.getElementsByTagName("nNF")
                        clienteOrder = xml.getElementsByTagName("xPed") or xml.getElementsByTagName("nItemPed")
                        time = xml.getElementsByTagName("dhEmi")
                        productCode = xml.getElementsByTagName("cProd")
                        amountOrdered = xml.getElementsByTagName("qCom")
                
#Utilizada para obter o número da nota fiscal do xml    
                for tag in nf:
                        nota = [(tag.firstChild.data)]

#Utilizada para obter o número ordem de compra da ALbacete do xml    
                for tag in clienteOrder:
#variável receve o valor encontrado no xml
                        clienteOrder = [(tag.firstChild.data)]

#Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                for tag in time:
                        pass  
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                        Datelist = [(tag.firstChild.data[0:4])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                        Datelist1 = [(tag.firstChild.data[5:7])]
# Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                        Datelist2 = [(tag.firstChild.data[8:10])]
# As listas são somadas na ordem desejada           
                        finalDate = Datelist2+Datelist1+Datelist
                        
#Converte lista para string, já colocando /
                        date = ["/".join(finalDate)]

#Utilizada para obter a referência do produto    
                for tag in productCode:
                        productCode = [(tag.firstChild.data[10:18])]


#Utilizada para obter a referência do produto    
                for tag in amountOrdered:
                        pass
                        amountOrdered = [(tag.firstChild.data[0])]
                        
                invoiceList = np.array([nota+clienteOrder+date+productCode+amountOrdered])
                print(invoiceList.flatten().tolist())
                pass
        
        else:
                print("Arquivo nao existe")
        weg = False
        break




#Usando a biblioteca OS é possível verificar se um arquivo xml existe e caso exista,
# o arquivo é excluído.


#While para apagar cada arquivo de XML
""""
while True:
        conditional = True
        if os.path.exists(wegInvoice):
           os.remove(wegInvoice)

        if os.path.exists(wegInvoice0):
             os.remove(wegInvoice0)
        
        if os.path.exists(wegInvoice1):
             os.remove(wegInvoice1)
        
        if os.path.exists(wegInvoice2):
             os.remove(wegInvoice2)
        
        if os.path.exists(wegInvoice3):
             os.remove(wegInvoice3)
        
        if os.path.exists(wegInvoice4):
             os.remove(wegInvoice4)
        
        if os.path.exists(wegInvoice5):
             os.remove(wegInvoice5)

        if os.path.exists(wegInvoice6):
             os.remove(wegInvoice6)

        if os.path.exists(wegInvoice7):
             os.remove(wegInvoice7)
        
        if os.path.exists(wegInvoice8):
             os.remove(wegInvoice8)
        
        if os.path.exists(wegInvoice9):
             os.remove(wegInvoice9)
        
        if os.path.exists(wegInvoice10):
             os.remove(wegInvoice10)
        
        if os.path.exists(wegInvoice11):
             os.remove(wegInvoice11)
        
        if os.path.exists(wegInvoice12):
             os.remove(wegInvoice12)

        if os.path.exists(wegInvoice13):
             os.remove(wegInvoice13)
        
        if os.path.exists(wegInvoice14):
             os.remove(wegInvoice14)
        
        if os.path.exists(wegInvoice15):
             os.remove(wegInvoice15)  
         
        else:
                print('Concluído')
        conditional = False
        break
"""