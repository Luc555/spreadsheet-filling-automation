import os.path
import subprocess
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from time import sleep
from tkinter import *
from tkinter import messagebox
from xml.dom import minidom

import gspread
import numpy as np
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

try:
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

#Computador casa
#searchPath = 'E:/Albacete-automation/Albacete-Automation/Automation-of-Spreadsheets\WEG-INVOICE'
        
#Computador trabalho
        searchPath = 'E:/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE'

        invoiceArray = []
        contagem = []

        for file in os.listdir("E:/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE"):
                if file.endswith(".xml"):
                        #invoiceArray = [os.path.join(file)]
                        invoiceArray.append(os.path.join(file))
                        print(searchPath+"/"+file)
                        print(len(invoiceArray))

                        pass
        
        print("\n")   
        print(invoiceArray)

        i = -1
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
                                        clienteOrder = [(tag.firstChild.data)]
                                        clienteOrderLoop.append(clienteOrder)
                                        for x in clienteOrderLoop:
                                                pass
                                
                                if clienteOrder == []:
                                        clienteOrder = ['Not exist']

                                elif clienteOrder == ['WMP - AMOSTRA']:
                                                
                                        clienteOrder = ['Sample']
                                else:
                                        pass

                        #Utilizada para obter a data de emissão da nota fiscal, onde é obtido o dado do xml(2022-05-18T07:46:31-03:00, por exemplo)    
                                for tag in time:
                                        pass  
                        # Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['2022'])        
                                        Datelist = [(tag.firstChild.data[0:4])]
                        # Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['05'])        
                                        Datelist1 = [(tag.firstChild.data[5:7])]
                                        if (Datelist1 == ['04']) or (Datelist1 == ['06']) or (Datelist1 == ['09']) or (Datelist1 == ['11']):
                                                totalMonthDays = 30
                                                pass
                                                
                                        elif (Datelist1 == ['01']) or (Datelist1 == ['03']) or (Datelist1 == ['05']) or (Datelist1 == ['07']) or (Datelist1 == ['08']) or (Datelist1 == ['10']) or (Datelist1 == ['12']):
                                                totalMonthDays = 31
                                                pass

                                        elif (Datelist1 == ['02']):
                                                totalMonthDays = 28
                                                pass
                        # Depois é criada uma lista e os dados do xml são colocadas dentro, buscando um determinado dado(['18'])        
                                        Datelist2 = [(tag.firstChild.data[8:10])]
                                        delivery = Datelist2[0]
                                        my_str = ''.join(delivery)
                                        delivery = int(my_str)
                                        delivery1 = delivery

                        # As listas são somadas na ordem desejada           
                                        finalDate = Datelist2+Datelist1+Datelist
                                                
                        #Converte lista para string, já colocando /
                                        date = ["/".join(finalDate)]
                        #Utilizada para obter a referência do produto
                                productCodeLoop = []
                                for tag in productCode:
                                        productCode = [(tag.firstChild.data[10:18])]
                                        productCodeLoop.append(productCode)
                #print('O total de índices é', len(productCodeLoop))

                # Se o código de referência for diferente de qualquer código listado a abaixo, valida a variável 'validate'
                # como verdadeira, que será chamada na frente.

                                if (productCode != ['14437060']) and (productCode != ['14437061']) and (productCode != ['14437062']) and (productCode != ['14437063']) and (productCode != ['14437064']) and (productCode != ['11432901']) and (productCode != ['11871633']) and (productCode != ['11873418']) and (productCode != ['14977774']) and (productCode != ['14977775']) and (productCode != ['14977776']) and (productCode != ['14977777']) and (productCode != ['14977938']) and (productCode != ['15079273']) and (productCode != ['14977939']) and (productCode != ['15308522']) and (productCode != ['15124776']) and (productCode != ['15083880']) and (productCode != ['15079268']):
                                        validate = True
                                        pass                      
                                else:
                #Transforma a lista referência(string) em uma lista de inteiros 
                                        if len(productCodeLoop) == 1:
                                                valores = productCodeLoop[0]
                                                ref = valores[0] # Primeiro valor da lista
                                                procv = [ planilha01.loc[planilha01['Ref.'] == int(ref), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref), 'Transportadora'].iloc[0]]
                                                code = [procv[0]]
                                                product_name = [procv[1]]
                                                delivery_day = [procv[2]]
                                                sum = (int(procv[2]))
                                                delivery1 = delivery1 + sum
                                                
                                        if len(productCodeLoop) == 2:
                #Aqui nós retiramos os valores presentes da lista em Loop de dentro de uma lista e deixamos na lista principal
                # a lista 'ref', por isso é usado o []+[]                                            
                                                ref = [productCodeLoop[0][0], productCodeLoop[1][0]] # Primeiro valor da lista
                #PAREI AQUI NO DIA 08/08/2022. VALIDANDO MAIS DE UMA ORDEM DE COMPRA EM UMA NOTA DA WEG                                                          
                                                procv = [ planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Transportadora'].iloc[0]]
                                                procv1 = [ planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Transportadora'].iloc[0]]                              
                                                code = [procv[0]]
                                                product_name = [procv[1]]
                                                code1 = [procv1[0]]
                                                product_name1 = [procv1[1]]
                                                
                                                
                                                delivery_day = [procv[2]]
                                                sum = (int(procv[2]))
                                                delivery1 = delivery1 + sum
                                                
                                        if len(productCodeLoop) == 3:
                        #Aqui nós retiramos os valores presentes da lista em Loop de dentro de uma lista e deixamos na lista principal
                # a lista 'ref', por isso é usado o []+[]                                            
                                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0]] # Primeiro valor da lista
                #PAREI AQUI NO DIA 08/08/2022. VALIDANDO MAIS DE UMA ORDEM DE COMPRA EM UMA NOTA DA WEG                                                          
                                                procv = [ planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Transportadora'].iloc[0]]
                                                procv1 = [ planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Transportadora'].iloc[0]]                              
                                                procv2 = [ planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Transportadora'].iloc[0]]                              
                                                code = [procv[0]]
                                                product_name = [procv[1]]
                                                code1 = [procv1[0]]
                                                product_name1 = [procv1[1]]
                                                code2 = [procv2[0]]
                                                product_name2 = [procv2[1]]
                                                
                                                
                                                delivery_day = [procv[2]]
                                                sum = (int(procv[2]))
                                                delivery1 = delivery1 + sum 
                                                
                                                
                                        if len(productCodeLoop) == 4:
                        #Aqui nós retiramos os valores presentes da lista em Loop de dentro de uma lista e deixamos na lista principal
                # a lista 'ref', por isso é usado o []+[]                                            
                                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0]] # Primeiro valor da lista
                #PAREI AQUI NO DIA 08/08/2022. VALIDANDO MAIS DE UMA ORDEM DE COMPRA EM UMA NOTA DA WEG                                                          
                                                procv = [ planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Transportadora'].iloc[0]]
                                                procv1 = [ planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Transportadora'].iloc[0]]                              
                                                procv2 = [ planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Transportadora'].iloc[0]]                              
                                                procv3 = [ planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Transportadora'].iloc[0]]                              


                                                code = [procv[0]]
                                                product_name = [procv[1]]
                                                code1 = [procv1[0]]
                                                product_name1 = [procv1[1]]
                                                code2 = [procv2[0]]
                                                product_name2 = [procv2[1]]
                                                code3 = [procv3[0]]
                                                product_name3 = [procv3[1]]
                                                
                                                
                                                delivery_day = [procv[2]]
                                                sum = (int(procv[2]))
                                                delivery1 = delivery1 + sum 
                                                
                                                
                                        
                                        if len(productCodeLoop) == 5:
                #Aqui nós retiramos os valores presentes da lista em Loop de dentro de uma lista e deixamos na lista principal
                # a lista 'ref', por isso é usado o []+[]                                            
                                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
                #PAREI AQUI NO DIA 08/08/2022. VALIDANDO MAIS DE UMA ORDEM DE COMPRA EM UMA NOTA DA WEG
                #Pensar em uma forma de fazer isso de forma automática, através de loop.                                                       
                                                procv = [ planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Transportadora'].iloc[0]]
                                                procv1 = [ planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Transportadora'].iloc[0]]                              
                                                procv2 = [ planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Transportadora'].iloc[0]]                              
                                                procv3 = [ planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Transportadora'].iloc[0]]                              
                                                procv4 = [ planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Transportadora'].iloc[0]]


                                                code = [procv[0]]
                                                product_name = [procv[1]]
                                                code1 = [procv1[0]]
                                                product_name1 = [procv1[1]]
                                                code2 = [procv2[0]]
                                                product_name2 = [procv2[1]]
                                                code3 = [procv3[0]]
                                                product_name3 = [procv3[1]]
                                                code4 = [procv4[0]]
                                                product_name4 = [procv4[1]]
                                                
                                                
                                                delivery_day = [procv[2]]
                                                sum = (int(procv[2]))
                                                delivery1 = delivery1 + sum 

                                                
                                                if delivery1 > totalMonthDays:
                                                        correctDay = delivery1- totalMonthDays
                                                        if correctDay<10:
                                                                delivery = str(correctDay)
                                                                delivery = '0' + delivery
                                                                delivery = [delivery]
                                                                month =Datelist1[0]
                                                                my_string = ''.join(month)
                                                                month = int(my_string)
                                                                month = month + 1
                                                        else:
                                                                delivery = str(correctDay)
                                                                delivery = [delivery]
                                                                month =Datelist1[0]
                                                                my_string = ''.join(month)
                                                                month = int(my_string)
                                                                month = month + 1
                                                        if month<10:
                                                                Datelist1 = str(month)
                                                                Datelist1 = '0' + Datelist1
                                                                Datelist1 = [Datelist1] 
                                                        else:       
                                                                Datelist1 = str(month)
                                                                Datelist1 = [Datelist1]
                                                                pass
                                                                

                                                        delivery = delivery + Datelist1 + Datelist
                                                        final_delivery = ["/".join(delivery)]
                                                        shipping_company = [procv[3]]
                                                        validate = False

                                                                
                                                else:
                                                        delivery = str(delivery1)
                                                        delivery = [delivery]
                                                        delivery = delivery + Datelist1 + Datelist
                                                        final_delivery = ["/".join(delivery)]
                                                        shipping_company = [procv[3]]
                                                        validate = False


                #Utilizada para obter a referência do produto    
                                        amountOrderedLoop = []
                                        for tag in amountOrdered:
                                                pass
                                                amountOrdered = (tag.firstChild.data[0:4])
                                                int_list = float(amountOrdered)
                                                amountOrdered = int(int_list)
                                                amountOrdered = [amountOrdered]
                                                amountOrderedLoop.append(amountOrdered)
                                                for x in productCodeLoop:
                                                        pass
                                                        
                                        if validate == True:
                                                invoiceZeroList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                                                invoiceZeroList=invoiceZeroList.flatten().tolist()
                                                print('invoice: ', invoiceZeroList)
                                                print('\n')
                                                
                                                worksheet.append_row(invoiceZeroList, value_input_option='USER_ENTERED')
                                        else:   
                                                if len(clienteOrderLoop) == 1:
                                                        invoiceZeroList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroList=invoiceZeroList.flatten().tolist()
                                                        print('invoice'+str(i)+': ',invoiceZeroList)
                                                        print('\n')
                                                                
                                                        worksheet.append_row(invoiceZeroList, value_input_option='USER_ENTERED')
                                                
                                                elif len(clienteOrderLoop) == 2:
                                                        invoiceZeroListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                                        invoiceZeroListLoopOne=invoiceZeroListLoopOne.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopOne)
                                                        print('\n')
                                                        invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopTwo)
                                                        print('\n')   
                                                        
                                                        worksheet.append_row(invoiceZeroListLoopOne, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopTwo, value_input_option='USER_ENTERED')
                                                        
                                                elif len(clienteOrderLoop) == 3:
                                                        invoiceZeroListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                                        invoiceZeroListLoopOne=invoiceZeroListLoopOne.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopOne)
                                                        print('\n')
                                                        invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopTwo)
                                                        print('\n')   
                                                        invoiceZeroListLoopThree=invoiceZeroListLoopThree.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopThree)
                                                        print('\n')
                                                        worksheet.append_row(invoiceZeroListLoopOne, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopTwo, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopThree, value_input_option='USER_ENTERED')

                                                elif len(clienteOrderLoop) == 4:
                                                        invoiceZeroListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                                        
                                                        invoiceZeroListLoopOne=invoiceZeroListLoopOne.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopOne)
                                                        print('\n')
                                                        invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopTwo)
                                                        print('\n')   
                                                        invoiceZeroListLoopThree=invoiceZeroListLoopThree.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopThree)
                                                        print('\n')
                                                        invoiceZeroListLoopFour=invoiceZeroListLoopFour.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopFour)
                                                        print('\n')
                                                        worksheet.append_row(invoiceZeroListLoopOne, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopTwo, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopThree, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopFour, value_input_option='USER_ENTERED')

                                                elif len(clienteOrderLoop) == 5:
                                                        invoiceZeroListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                                        invoiceZeroListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                                        
                                                        invoiceZeroListLoopOne=invoiceZeroListLoopOne.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopOne)
                                                        print('\n')
                                                        invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopTwo)
                                                        print('\n')   
                                                        invoiceZeroListLoopThree=invoiceZeroListLoopThree.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopThree)
                                                        print('\n')
                                                        invoiceZeroListLoopFour=invoiceZeroListLoopFour.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopFour)
                                                        print('\n')
                                                        invoiceZeroListLoopFive=invoiceZeroListLoopFive.flatten().tolist()
                                                        print( 'invoice'+str(i)+': ',invoiceZeroListLoopFive)
                                                        print('\n')
                                                        worksheet.append_row(invoiceZeroListLoopOne, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopTwo, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopThree, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopFour, value_input_option='USER_ENTERED')
                                                        worksheet.append_row(invoiceZeroListLoopFive, value_input_option='USER_ENTERED')

                                pass
                        

                else:
                        print("Arquivo nao existe")
                weg = False
                break
        
except FileNotFoundError as error:
      vmsg = "Não encontramos nem o arquivo e nem o diretório"
      tiposmg = error
      def showMessage(tiposmg, msg):
            if tiposmg == error:
                  messagebox.showerror(title="Sem arquivos ou diretório", message=msg)

      showMessage(tiposmg, vmsg )
      
except WebDriverException as erro:       
      vmsg = "A versão do chromedriver está desatualizada.\n\n Favor atualizar a versão para que esta se adeque à versão do navegador:"
      print(type(vmsg))
      tiposmg = erro
      def showMessage(tiposmg, msg):
            batata = str(erro)
            delimiter =  ''
            batataList = delimiter.join(batata)
            takingWhatMatters = batataList[118:128]
            print(takingWhatMatters)
                  
            if tiposmg == erro:
                  messagebox.showerror(title="Chromedriver desatualizado", message=msg+takingWhatMatters+".\n\n Este é o link: https://chromedriver.chromium.org/downloads")
                  print("Erro - exceção Webdriverexception: ", erro)
      showMessage(tiposmg, vmsg)