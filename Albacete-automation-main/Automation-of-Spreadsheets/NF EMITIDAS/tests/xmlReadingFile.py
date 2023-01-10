from xml.dom import minidom
from time import sleep
import os
import os.path
import numpy as np
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials


#Com essa função é possível rodar o script que abre os arquivos xml e buscar as informações dentro de cada  
# nota fiscal.



#Escopo utilizado
scope = ['https://spreadsheets.google.com/feeds']

#Dados de autenticação
credentials = ServiceAccountCredentials.from_json_keyfile_name('E:/Albacete-automation/Credential_google.json', scope)

#Se autentica
gc = gspread.authorize(credentials)


#Abre a planilha
wks = gc.open_by_key('1Sy5HQBbSRewZrr3ZLtCaCHRqVul2ihajVoLscePMpVw')

#Para selecionar a planilha pelo o nome use o código abaixo
#wks = gc.open('Teste Python') 

#Seleciona a primeira página da planilha
worksheet = wks.get_worksheet(15)


#Código que lê a planilha referenciada e sua respectiva aba
planilha01 = pd.read_excel("E:/Albacete-automation/DATABASE/Parametros-dos-motores.xlsx", sheet_name="Parâmetros dos Motores")


wegInvoice = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice.xml'
wegInvoice0 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice0.xml'
wegInvoice1 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice1.xml'
wegInvoice2 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice2.xml'
wegInvoice3 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice3.xml'
wegInvoice4 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice4.xml'
wegInvoice5 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice5.xml'
wegInvoice6 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice6.xml'
wegInvoice7 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice7.xml'
wegInvoice8 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice8.xml'
wegInvoice9 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice9.xml'
wegInvoice10 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice10.xml'
wegInvoice11 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice11.xml'
wegInvoice12 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice12.xml'
wegInvoice13 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice13.xml'
wegInvoice14 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice14.xml'
wegInvoice15 = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice/wegInvoice15.xml'




searchPath = 'E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice'

for file in os.listdir("E:/Albacete-automation/Automation-of-Spreadsheets/weg-invoice"):
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
                validation=[]
                if (productCode != ['14437060']) and (productCode != ['14437061']) and (productCode != ['14437062']) and (productCode != ['14437063']) and (productCode != ['14437064']) and (productCode != ['11432901']) and (productCode != ['11871633']) and (productCode != ['11873418']) and (productCode != ['14977774']) and (productCode != ['14977775']) and (productCode != ['14977776']) and (productCode != ['14977777']) and (productCode != ['14977938']) and (productCode != ['15079273']) and (productCode != ['14977939']) and (productCode != ['15308522']) and (productCode != ['15124776']) and (productCode != ['15083880']) and (productCode != ['15079268']):
                        validation.append(True)
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
                                pass
                                
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
                str(validation)
                print(validation)
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
                                
                if validation == True:
                        invoiceZeroList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceZeroList=invoiceZeroList.flatten().tolist()
                        print('\n')
                        
                        worksheet.append_row(invoiceZeroList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceZeroList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceZeroList=invoiceZeroList.flatten().tolist()
                                print('\n')
                                        
                                worksheet.append_row(invoiceZeroList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceZeroListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceZeroListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceZeroListLoopOne=invoiceZeroListLoopOne.flatten().tolist()
                                print('\n')
                                invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                print('\n')   
                                
                                worksheet.append_row(invoiceZeroListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceZeroListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceZeroListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceZeroListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceZeroListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceZeroListLoopOne=invoiceZeroListLoopOne.flatten().tolist()
                                print('\n')
                                invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                print('\n')   
                                invoiceZeroListLoopThree=invoiceZeroListLoopThree.flatten().tolist()
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
                                print('\n')
                                invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                print('\n')   
                                invoiceZeroListLoopThree=invoiceZeroListLoopThree.flatten().tolist()
                                print('\n')
                                invoiceZeroListLoopFour=invoiceZeroListLoopFour.flatten().tolist()
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
                                print('\n')
                                invoiceZeroListLoopTwo=invoiceZeroListLoopTwo.flatten().tolist()
                                print('\n')   
                                invoiceZeroListLoopThree=invoiceZeroListLoopThree.flatten().tolist()
                                print('\n')
                                invoiceZeroListLoopFour=invoiceZeroListLoopFour.flatten().tolist()
                                print('\n')
                                invoiceZeroListLoopFive=invoiceZeroListLoopFive.flatten().tolist()
                                print('\n')
                                worksheet.append_row(invoiceZeroListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceZeroListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceZeroListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceZeroListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceZeroListLoopFive, value_input_option='USER_ENTERED')

                        pass
                 
        if os.path.exists(wegInvoice1):
                with open(wegInvoice1, 'r', encoding='utf-8') as f:
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
                                pass
                                
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
                        invoiceOneList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceOneList=invoiceOneList.flatten().tolist()
                        print('invoice1: ', invoiceOneList)
                        print('\n')
                        
                        worksheet.append_row(invoiceOneList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceOneList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceOneList=invoiceOneList.flatten().tolist()
                                print( 'invoice1: ',invoiceOneList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceOneList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceOneListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceOneListLoopOne=invoiceOneListLoopOne.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopOne)
                                print('\n')
                                invoiceOneListLoopTwo=invoiceOneListLoopTwo.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceOneListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceOneListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceOneListLoopOne=invoiceOneListLoopOne.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopOne)
                                print('\n')
                                invoiceOneListLoopTwo=invoiceOneListLoopTwo.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopTwo)
                                print('\n')   
                                invoiceOneListLoopThree=invoiceOneListLoopThree.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceOneListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceOneListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceOneListLoopOne=invoiceOneListLoopOne.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopOne)
                                print('\n')
                                invoiceOneListLoopTwo=invoiceOneListLoopTwo.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopTwo)
                                print('\n')   
                                invoiceOneListLoopThree=invoiceOneListLoopThree.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopThree)
                                print('\n')
                                invoiceOneListLoopFour=invoiceOneListLoopFour.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceOneListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceOneListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceOneListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceOneListLoopOne=invoiceOneListLoopOne.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopOne)
                                print('\n')
                                invoiceOneListLoopTwo=invoiceOneListLoopTwo.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopTwo)
                                print('\n')   
                                invoiceOneListLoopThree=invoiceOneListLoopThree.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopThree)
                                print('\n')
                                invoiceOneListLoopFour=invoiceOneListLoopFour.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopFour)
                                print('\n')
                                invoiceOneListLoopFive=invoiceOneListLoopFive.flatten().tolist()
                                print( 'invoice1: ',invoiceOneListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceOneListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceOneListLoopFive, value_input_option='USER_ENTERED')

                        pass
        
        if os.path.exists(wegInvoice2):
                with open(wegInvoice2, 'r', encoding='utf-8') as f:
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
                        invoiceTwoList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceTwoList=invoiceTwoList.flatten().tolist()
                        print('invoice2: ', invoiceTwoList)
                        print('\n')
                        
                        worksheet.append_row(invoiceTwoList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceTwoList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceTwoList=invoiceTwoList.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceTwoList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceTwoListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceTwoListLoopOne=invoiceTwoListLoopOne.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopOne)
                                print('\n')
                                invoiceTwoListLoopTwo=invoiceTwoListLoopTwo.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceTwoListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceTwoListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceTwoListLoopOne=invoiceTwoListLoopOne.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopOne)
                                print('\n')
                                invoiceTwoListLoopTwo=invoiceTwoListLoopTwo.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopTwo)
                                print('\n')   
                                invoiceTwoListLoopThree=invoiceTwoListLoopThree.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceTwoListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceTwoListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceTwoListLoopOne=invoiceTwoListLoopOne.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopOne)
                                print('\n')
                                invoiceTwoListLoopTwo=invoiceTwoListLoopTwo.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopTwo)
                                print('\n')   
                                invoiceTwoListLoopThree=invoiceTwoListLoopThree.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopThree)
                                print('\n')
                                invoiceTwoListLoopFour=invoiceTwoListLoopFour.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceTwoListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceTwoListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceTwoListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceTwoListLoopOne=invoiceTwoListLoopOne.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopOne)
                                print('\n')
                                invoiceTwoListLoopTwo=invoiceTwoListLoopTwo.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopTwo)
                                print('\n')   
                                invoiceTwoListLoopThree=invoiceTwoListLoopThree.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopThree)
                                print('\n')
                                invoiceTwoListLoopFour=invoiceTwoListLoopFour.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopFour)
                                print('\n')
                                invoiceTwoListLoopFive=invoiceTwoListLoopFive.flatten().tolist()
                                print( 'invoice2: ',invoiceTwoListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceTwoListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwoListLoopFive, value_input_option='USER_ENTERED')

                        pass                
        
        if os.path.exists(wegInvoice3):
                with open(wegInvoice3, 'r', encoding='utf-8') as f:
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
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
                        invoiceThreeList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceThreeList=invoiceThreeList.flatten().tolist()
                        print('invoice3: ', invoiceThreeList)
                        print('\n')
                        
                        worksheet.append_row(invoiceThreeList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceThreeList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceThreeList=invoiceThreeList.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceThreeList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceThreeListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceThreeListLoopOne=invoiceThreeListLoopOne.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopOne)
                                print('\n')
                                invoiceThreeListLoopTwo=invoiceThreeListLoopTwo.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceThreeListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceThreeListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceThreeListLoopOne=invoiceThreeListLoopOne.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopOne)
                                print('\n')
                                invoiceThreeListLoopTwo=invoiceThreeListLoopTwo.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopTwo)
                                print('\n')   
                                invoiceThreeListLoopThree=invoiceThreeListLoopThree.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceThreeListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceThreeListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceThreeListLoopOne=invoiceThreeListLoopOne.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopOne)
                                print('\n')
                                invoiceThreeListLoopTwo=invoiceThreeListLoopTwo.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopTwo)
                                print('\n')   
                                invoiceThreeListLoopThree=invoiceThreeListLoopThree.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopThree)
                                print('\n')
                                invoiceThreeListLoopFour=invoiceThreeListLoopFour.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceThreeListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceThreeListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceThreeListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceThreeListLoopOne=invoiceThreeListLoopOne.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopOne)
                                print('\n')
                                invoiceThreeListLoopTwo=invoiceThreeListLoopTwo.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopTwo)
                                print('\n')   
                                invoiceThreeListLoopThree=invoiceThreeListLoopThree.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopThree)
                                print('\n')
                                invoiceThreeListLoopFour=invoiceThreeListLoopFour.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopFour)
                                print('\n')
                                invoiceThreeListLoopFive=invoiceThreeListLoopFive.flatten().tolist()
                                print( 'invoice3: ',invoiceThreeListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceThreeListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThreeListLoopFive, value_input_option='USER_ENTERED')

                        pass                

        if os.path.exists(wegInvoice4):
                with open(wegInvoice4, 'r', encoding='utf-8') as f:
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0] ] # Primeiro valor da lista
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
                        invoiceFourList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceFourList=invoiceFourList.flatten().tolist()
                        print('invoice4: ', invoiceFourList)
                        print('\n')
                        
                        worksheet.append_row(invoiceFourList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceFourList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceFourList=invoiceFourList.flatten().tolist()
                                print( 'invoice4: ',invoiceFourList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceFourList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceFourListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceFourListLoopOne=invoiceFourListLoopOne.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopOne)
                                print('\n')
                                invoiceFourListLoopTwo=invoiceFourListLoopTwo.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceFourListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceFourListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceFourListLoopOne=invoiceFourListLoopOne.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopOne)
                                print('\n')
                                invoiceFourListLoopTwo=invoiceFourListLoopTwo.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopTwo)
                                print('\n')   
                                invoiceFourListLoopThree=invoiceFourListLoopThree.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceFourListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceFourListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFourListLoopOne=invoiceFourListLoopOne.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopOne)
                                print('\n')
                                invoiceFourListLoopTwo=invoiceFourListLoopTwo.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopTwo)
                                print('\n')   
                                invoiceFourListLoopThree=invoiceFourListLoopThree.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopThree)
                                print('\n')
                                invoiceFourListLoopFour=invoiceFourListLoopFour.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceFourListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceFourListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceFourListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFourListLoopOne=invoiceFourListLoopOne.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopOne)
                                print('\n')
                                invoiceFourListLoopTwo=invoiceFourListLoopTwo.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopTwo)
                                print('\n')   
                                invoiceFourListLoopThree=invoiceFourListLoopThree.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopThree)
                                print('\n')
                                invoiceFourListLoopFour=invoiceFourListLoopFour.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopFour)
                                print('\n')
                                invoiceFourListLoopFive=invoiceFourListLoopFive.flatten().tolist()
                                print( 'invoice4: ',invoiceFourListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceFourListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourListLoopFive, value_input_option='USER_ENTERED')

                        pass                
        if os.path.exists(wegInvoice5):
                with open(wegInvoice5, 'r', encoding='utf-8') as f:
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
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
                        invoiceFiveList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceFiveList=invoiceFiveList.flatten().tolist()
                        print('invoice5: ', invoiceFiveList)
                        print('\n')
                        
                        worksheet.append_row(invoiceFiveList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceFiveList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceFiveList=invoiceFiveList.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceFiveList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceFiveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceFiveListLoopOne=invoiceFiveListLoopOne.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopOne)
                                print('\n')
                                invoiceFiveListLoopTwo=invoiceFiveListLoopTwo.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceFiveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceFiveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceFiveListLoopOne=invoiceFiveListLoopOne.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopOne)
                                print('\n')
                                invoiceFiveListLoopTwo=invoiceFiveListLoopTwo.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopTwo)
                                print('\n')   
                                invoiceFiveListLoopThree=invoiceFiveListLoopThree.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceFiveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceFiveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFiveListLoopOne=invoiceFiveListLoopOne.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopOne)
                                print('\n')
                                invoiceFiveListLoopTwo=invoiceFiveListLoopTwo.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopTwo)
                                print('\n')   
                                invoiceFiveListLoopThree=invoiceFiveListLoopThree.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopThree)
                                print('\n')
                                invoiceFiveListLoopFour=invoiceFiveListLoopFour.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceFiveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceFiveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceFiveListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFiveListLoopOne=invoiceFiveListLoopOne.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopOne)
                                print('\n')
                                invoiceFiveListLoopTwo=invoiceFiveListLoopTwo.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopTwo)
                                print('\n')   
                                invoiceFiveListLoopThree=invoiceFiveListLoopThree.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopThree)
                                print('\n')
                                invoiceFiveListLoopFour=invoiceFiveListLoopFour.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopFour)
                                print('\n')
                                invoiceFiveListLoopFive=invoiceFiveListLoopFive.flatten().tolist()
                                print( 'invoice5: ',invoiceFiveListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceFiveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFiveListLoopFive, value_input_option='USER_ENTERED')

                        pass
        if os.path.exists(wegInvoice6):
                with open(wegInvoice6, 'r', encoding='utf-8') as f:
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
                        invoiceSixList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceSixList=invoiceSixList.flatten().tolist()
                        print('invoice6: ', invoiceSixList)
                        print('\n')
                        
                        worksheet.append_row(invoiceSixList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceSixList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceSixList=invoiceSixList.flatten().tolist()
                                print( 'invoice6: ',invoiceSixList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceSixList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceSixListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceSixListLoopOne=invoiceSixListLoopOne.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopOne)
                                print('\n')
                                invoiceSixListLoopTwo=invoiceSixListLoopTwo.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceSixListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceSixListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceSixListLoopOne=invoiceSixListLoopOne.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopOne)
                                print('\n')
                                invoiceSixListLoopTwo=invoiceSixListLoopTwo.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopTwo)
                                print('\n')   
                                invoiceSixListLoopThree=invoiceSixListLoopThree.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceSixListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceSixListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceSixListLoopOne=invoiceSixListLoopOne.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopOne)
                                print('\n')
                                invoiceSixListLoopTwo=invoiceSixListLoopTwo.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopTwo)
                                print('\n')   
                                invoiceSixListLoopThree=invoiceSixListLoopThree.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopThree)
                                print('\n')
                                invoiceSixListLoopFour=invoiceSixListLoopFour.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceSixListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceSixListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceSixListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceSixListLoopOne=invoiceSixListLoopOne.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopOne)
                                print('\n')
                                invoiceSixListLoopTwo=invoiceSixListLoopTwo.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopTwo)
                                print('\n')   
                                invoiceSixListLoopThree=invoiceSixListLoopThree.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopThree)
                                print('\n')
                                invoiceSixListLoopFour=invoiceSixListLoopFour.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopFour)
                                print('\n')
                                invoiceSixListLoopFive=invoiceSixListLoopFive.flatten().tolist()
                                print( 'invoice6: ',invoiceSixListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceSixListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSixListLoopFive, value_input_option='USER_ENTERED')

                        pass                
        if os.path.exists(wegInvoice7):
                with open(wegInvoice7, 'r', encoding='utf-8') as f:
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
                        invoiceSevenList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceSevenList=invoiceSevenList.flatten().tolist()
                        print('invoice7: ', invoiceSevenList)
                        print('\n')
                        
                        worksheet.append_row(invoiceSevenList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceSevenList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceSevenList=invoiceSevenList.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceSevenList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceSevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceSevenListLoopOne=invoiceSevenListLoopOne.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopOne)
                                print('\n')
                                invoiceSevenListLoopTwo=invoiceSevenListLoopTwo.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceSevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceSevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceSevenListLoopOne=invoiceSevenListLoopOne.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopOne)
                                print('\n')
                                invoiceSevenListLoopTwo=invoiceSevenListLoopTwo.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopTwo)
                                print('\n')   
                                invoiceSevenListLoopThree=invoiceSevenListLoopThree.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceSevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceSevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceSevenListLoopOne=invoiceSevenListLoopOne.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopOne)
                                print('\n')
                                invoiceSevenListLoopTwo=invoiceSevenListLoopTwo.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopTwo)
                                print('\n')   
                                invoiceSevenListLoopThree=invoiceSevenListLoopThree.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopThree)
                                print('\n')
                                invoiceSevenListLoopFour=invoiceSevenListLoopFour.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceSevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceSevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceSevenListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceSevenListLoopOne=invoiceSevenListLoopOne.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopOne)
                                print('\n')
                                invoiceSevenListLoopTwo=invoiceSevenListLoopTwo.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopTwo)
                                print('\n')   
                                invoiceSevenListLoopThree=invoiceSevenListLoopThree.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopThree)
                                print('\n')
                                invoiceSevenListLoopFour=invoiceSevenListLoopFour.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopFour)
                                print('\n')
                                invoiceSevenListLoopFive=invoiceSevenListLoopFive.flatten().tolist()
                                print( 'invoice7: ',invoiceSevenListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceSevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceSevenListLoopFive, value_input_option='USER_ENTERED')

                        pass       
        if os.path.exists(wegInvoice8):
                with open(wegInvoice8, 'r', encoding='utf-8') as f:
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
                        invoiceEightList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceEightList=invoiceEightList.flatten().tolist()
                        print('invoice8: ', invoiceEightList)
                        print('\n')
                        
                        worksheet.append_row(invoiceEightList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceEightList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceEightList=invoiceEightList.flatten().tolist()
                                print( 'invoice8: ',invoiceEightList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceEightList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceEightListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceEightListLoopOne=invoiceEightListLoopOne.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopOne)
                                print('\n')
                                invoiceEightListLoopTwo=invoiceEightListLoopTwo.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceEightListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceEightListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceEightListLoopOne=invoiceEightListLoopOne.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopOne)
                                print('\n')
                                invoiceEightListLoopTwo=invoiceEightListLoopTwo.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopTwo)
                                print('\n')   
                                invoiceEightListLoopThree=invoiceEightListLoopThree.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceEightListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceEightListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceEightListLoopOne=invoiceEightListLoopOne.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopOne)
                                print('\n')
                                invoiceEightListLoopTwo=invoiceEightListLoopTwo.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopTwo)
                                print('\n')   
                                invoiceEightListLoopThree=invoiceEightListLoopThree.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopThree)
                                print('\n')
                                invoiceEightListLoopFour=invoiceEightListLoopFour.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceEightListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceEightListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceEightListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceEightListLoopOne=invoiceEightListLoopOne.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopOne)
                                print('\n')
                                invoiceEightListLoopTwo=invoiceEightListLoopTwo.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopTwo)
                                print('\n')   
                                invoiceEightListLoopThree=invoiceEightListLoopThree.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopThree)
                                print('\n')
                                invoiceEightListLoopFour=invoiceEightListLoopFour.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopFour)
                                print('\n')
                                invoiceEightListLoopFive=invoiceEightListLoopFive.flatten().tolist()
                                print( 'invoice8: ',invoiceEightListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceEightListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceEightListLoopFive, value_input_option='USER_ENTERED')

                        pass                        
        if os.path.exists(wegInvoice9):
                with open(wegInvoice9, 'r', encoding='utf-8') as f:
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
                        invoiceNineList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceNineList=invoiceNineList.flatten().tolist()
                        print('invoice9: ', invoiceNineList)
                        print('\n')
                        
                        worksheet.append_row(invoiceNineList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceNineList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceNineList=invoiceNineList.flatten().tolist()
                                print( 'invoice9: ',invoiceNineList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceNineList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceNineListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceNineListLoopOne=invoiceNineListLoopOne.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopOne)
                                print('\n')
                                invoiceNineListLoopTwo=invoiceNineListLoopTwo.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceNineListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceNineListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceNineListLoopOne=invoiceNineListLoopOne.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopOne)
                                print('\n')
                                invoiceNineListLoopTwo=invoiceNineListLoopTwo.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopTwo)
                                print('\n')   
                                invoiceNineListLoopThree=invoiceNineListLoopThree.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceNineListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceNineListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceNineListLoopOne=invoiceNineListLoopOne.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopOne)
                                print('\n')
                                invoiceNineListLoopTwo=invoiceNineListLoopTwo.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopTwo)
                                print('\n')   
                                invoiceNineListLoopThree=invoiceNineListLoopThree.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopThree)
                                print('\n')
                                invoiceNineListLoopFour=invoiceNineListLoopFour.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceNineListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceNineListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceNineListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceNineListLoopOne=invoiceNineListLoopOne.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopOne)
                                print('\n')
                                invoiceNineListLoopTwo=invoiceNineListLoopTwo.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopTwo)
                                print('\n')   
                                invoiceNineListLoopThree=invoiceNineListLoopThree.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopThree)
                                print('\n')
                                invoiceNineListLoopFour=invoiceNineListLoopFour.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopFour)
                                print('\n')
                                invoiceNineListLoopFive=invoiceNineListLoopFive.flatten().tolist()
                                print( 'invoice9: ',invoiceNineListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceNineListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceNineListLoopFive, value_input_option='USER_ENTERED')

                        pass                
        if os.path.exists(wegInvoice10):
                with open(wegInvoice10, 'r', encoding='utf-8') as f:
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
                        invoiceTenList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceTenList=invoiceTenList.flatten().tolist()
                        print('invoice10: ', invoiceTenList)
                        print('\n')
                        
                        worksheet.append_row(invoiceTenList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceTenList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceTenList=invoiceTenList.flatten().tolist()
                                print( 'invoice10: ',invoiceTenList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceTenList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceTenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceTenListLoopOne=invoiceTenListLoopOne.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopOne)
                                print('\n')
                                invoiceTenListLoopTwo=invoiceTenListLoopTwo.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceTenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceTenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceTenListLoopOne=invoiceTenListLoopOne.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopOne)
                                print('\n')
                                invoiceTenListLoopTwo=invoiceTenListLoopTwo.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopTwo)
                                print('\n')   
                                invoiceTenListLoopThree=invoiceTenListLoopThree.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceTenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceTenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceTenListLoopOne=invoiceTenListLoopOne.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopOne)
                                print('\n')
                                invoiceTenListLoopTwo=invoiceTenListLoopTwo.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopTwo)
                                print('\n')   
                                invoiceTenListLoopThree=invoiceTenListLoopThree.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopThree)
                                print('\n')
                                invoiceTenListLoopFour=invoiceTenListLoopFour.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceTenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceTenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceTenListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceTenListLoopOne=invoiceTenListLoopOne.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopOne)
                                print('\n')
                                invoiceTenListLoopTwo=invoiceTenListLoopTwo.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopTwo)
                                print('\n')   
                                invoiceTenListLoopThree=invoiceTenListLoopThree.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopThree)
                                print('\n')
                                invoiceTenListLoopFour=invoiceTenListLoopFour.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopFour)
                                print('\n')
                                invoiceTenListLoopFive=invoiceTenListLoopFive.flatten().tolist()
                                print( 'invoice10: ',invoiceTenListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceTenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTenListLoopFive, value_input_option='USER_ENTERED')

                        pass                
        if os.path.exists(wegInvoice11):
                with open(wegInvoice11, 'r', encoding='utf-8') as f:
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
                        invoiceElevenList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceElevenList=invoiceElevenList.flatten().tolist()
                        print('invoice11: ', invoiceElevenList)
                        print('\n')
                        
                        worksheet.append_row(invoiceElevenList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceElevenList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceElevenList=invoiceElevenList.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceElevenList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceElevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceElevenListLoopOne=invoiceElevenListLoopOne.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopOne)
                                print('\n')
                                invoiceElevenListLoopTwo=invoiceElevenListLoopTwo.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceElevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceElevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceElevenListLoopOne=invoiceElevenListLoopOne.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopOne)
                                print('\n')
                                invoiceElevenListLoopTwo=invoiceElevenListLoopTwo.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopTwo)
                                print('\n')   
                                invoiceElevenListLoopThree=invoiceElevenListLoopThree.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceElevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceElevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceElevenListLoopOne=invoiceElevenListLoopOne.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopOne)
                                print('\n')
                                invoiceElevenListLoopTwo=invoiceElevenListLoopTwo.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopTwo)
                                print('\n')   
                                invoiceElevenListLoopThree=invoiceElevenListLoopThree.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopThree)
                                print('\n')
                                invoiceElevenListLoopFour=invoiceElevenListLoopFour.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceElevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceElevenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceElevenListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceElevenListLoopOne=invoiceElevenListLoopOne.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopOne)
                                print('\n')
                                invoiceElevenListLoopTwo=invoiceElevenListLoopTwo.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopTwo)
                                print('\n')   
                                invoiceElevenListLoopThree=invoiceElevenListLoopThree.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopThree)
                                print('\n')
                                invoiceElevenListLoopFour=invoiceElevenListLoopFour.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopFour)
                                print('\n')
                                invoiceElevenListLoopFive=invoiceElevenListLoopFive.flatten().tolist()
                                print( 'invoice11: ',invoiceElevenListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceElevenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceElevenListLoopFive, value_input_option='USER_ENTERED')

                        pass
        
        if os.path.exists(wegInvoice12):
                with open(wegInvoice12, 'r', encoding='utf-8') as f:
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
                        invoiceTwelveList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceTwelveList=invoiceTwelveList.flatten().tolist()
                        print('invoice12: ', invoiceTwelveList)
                        print('\n')
                        
                        worksheet.append_row(invoiceTwelveList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceTwelveList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceTwelveList=invoiceTwelveList.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceTwelveList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceTwelveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceTwelveListLoopOne=invoiceTwelveListLoopOne.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopOne)
                                print('\n')
                                invoiceTwelveListLoopTwo=invoiceTwelveListLoopTwo.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceTwelveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceTwelveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceTwelveListLoopOne=invoiceTwelveListLoopOne.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopOne)
                                print('\n')
                                invoiceTwelveListLoopTwo=invoiceTwelveListLoopTwo.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopTwo)
                                print('\n')   
                                invoiceTwelveListLoopThree=invoiceTwelveListLoopThree.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceTwelveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceTwelveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceTwelveListLoopOne=invoiceTwelveListLoopOne.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopOne)
                                print('\n')
                                invoiceTwelveListLoopTwo=invoiceTwelveListLoopTwo.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopTwo)
                                print('\n')   
                                invoiceTwelveListLoopThree=invoiceTwelveListLoopThree.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopThree)
                                print('\n')
                                invoiceTwelveListLoopFour=invoiceTwelveListLoopFour.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceTwelveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceTwelveListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceTwelveListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceTwelveListLoopOne=invoiceTwelveListLoopOne.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopOne)
                                print('\n')
                                invoiceTwelveListLoopTwo=invoiceTwelveListLoopTwo.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopTwo)
                                print('\n')   
                                invoiceTwelveListLoopThree=invoiceTwelveListLoopThree.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopThree)
                                print('\n')
                                invoiceTwelveListLoopFour=invoiceTwelveListLoopFour.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopFour)
                                print('\n')
                                invoiceTwelveListLoopFive=invoiceTwelveListLoopFive.flatten().tolist()
                                print( 'invoice12: ',invoiceTwelveListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceTwelveListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceTwelveListLoopFive, value_input_option='USER_ENTERED')

                        pass
                 
        if os.path.exists(wegInvoice13):
                with open(wegInvoice13, 'r', encoding='utf-8') as f:
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
                        invoiceThirteenList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceThirteenList=invoiceThirteenList.flatten().tolist()
                        print('invoice13: ', invoiceThirteenList)
                        print('\n')
                        
                        worksheet.append_row(invoiceThirteenList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceThirteenList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceThirteenList=invoiceThirteenList.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceThirteenList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceThirteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceThirteenListLoopOne=invoiceThirteenListLoopOne.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopOne)
                                print('\n')
                                invoiceThirteenListLoopTwo=invoiceThirteenListLoopTwo.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceThirteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceThirteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceThirteenListLoopOne=invoiceThirteenListLoopOne.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopOne)
                                print('\n')
                                invoiceThirteenListLoopTwo=invoiceThirteenListLoopTwo.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopTwo)
                                print('\n')   
                                invoiceThirteenListLoopThree=invoiceThirteenListLoopThree.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceThirteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceThirteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceThirteenListLoopOne=invoiceThirteenListLoopOne.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopOne)
                                print('\n')
                                invoiceThirteenListLoopTwo=invoiceThirteenListLoopTwo.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopTwo)
                                print('\n')   
                                invoiceThirteenListLoopThree=invoiceThirteenListLoopThree.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopThree)
                                print('\n')
                                invoiceThirteenListLoopFour=invoiceThirteenListLoopFour.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceThirteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceThirteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceThirteenListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceThirteenListLoopOne=invoiceThirteenListLoopOne.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopOne)
                                print('\n')
                                invoiceThirteenListLoopTwo=invoiceThirteenListLoopTwo.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopTwo)
                                print('\n')   
                                invoiceThirteenListLoopThree=invoiceThirteenListLoopThree.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopThree)
                                print('\n')
                                invoiceThirteenListLoopFour=invoiceThirteenListLoopFour.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopFour)
                                print('\n')
                                invoiceThirteenListLoopFive=invoiceThirteenListLoopFive.flatten().tolist()
                                print( 'invoice13: ',invoiceThirteenListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceThirteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceThirteenListLoopFive, value_input_option='USER_ENTERED')

                        pass
        
        if os.path.exists(wegInvoice14):
                with open(wegInvoice14, 'r', encoding='utf-8') as f:
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
                        invoiceFourteenList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceFourteenList=invoiceFourteenList.flatten().tolist()
                        print('invoice14: ', invoiceFourteenList)
                        print('\n')
                        
                        worksheet.append_row(invoiceFourteenList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceFourteenList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceFourteenList=invoiceFourteenList.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceFourteenList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceFourteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceFourteenListLoopOne=invoiceFourteenListLoopOne.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopOne)
                                print('\n')
                                invoiceFourteenListLoopTwo=invoiceFourteenListLoopTwo.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceFourteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceFourteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceFourteenListLoopOne=invoiceFourteenListLoopOne.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopOne)
                                print('\n')
                                invoiceFourteenListLoopTwo=invoiceFourteenListLoopTwo.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopTwo)
                                print('\n')   
                                invoiceFourteenListLoopThree=invoiceFourteenListLoopThree.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceFourteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceFourteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFourteenListLoopOne=invoiceFourteenListLoopOne.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopOne)
                                print('\n')
                                invoiceFourteenListLoopTwo=invoiceFourteenListLoopTwo.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopTwo)
                                print('\n')   
                                invoiceFourteenListLoopThree=invoiceFourteenListLoopThree.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopThree)
                                print('\n')
                                invoiceFourteenListLoopFour=invoiceFourteenListLoopFour.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceFourteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceFourteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceFourteenListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFourteenListLoopOne=invoiceFourteenListLoopOne.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopOne)
                                print('\n')
                                invoiceFourteenListLoopTwo=invoiceFourteenListLoopTwo.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopTwo)
                                print('\n')   
                                invoiceFourteenListLoopThree=invoiceFourteenListLoopThree.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopThree)
                                print('\n')
                                invoiceFourteenListLoopFour=invoiceFourteenListLoopFour.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopFour)
                                print('\n')
                                invoiceFourteenListLoopFive=invoiceFourteenListLoopFive.flatten().tolist()
                                print( 'invoice14: ',invoiceFourteenListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceFourteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFourteenListLoopFive, value_input_option='USER_ENTERED')

                        pass        
        
        if os.path.exists(wegInvoice15):
                with open(wegInvoice15, 'r', encoding='utf-8') as f:
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
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
                        invoiceFifteenList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceFifteenList=invoiceFifteenList.flatten().tolist()
                        print('invoice15: ', invoiceFifteenList)
                        print('\n')
                        
                        worksheet.append_row(invoiceFifteenList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceFifteenList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceFifteenList=invoiceFifteenList.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceFifteenList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceFifteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceFifteenListLoopOne=invoiceFifteenListLoopOne.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopOne)
                                print('\n')
                                invoiceFifteenListLoopTwo=invoiceFifteenListLoopTwo.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceFifteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceFifteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceFifteenListLoopOne=invoiceFifteenListLoopOne.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopOne)
                                print('\n')
                                invoiceFifteenListLoopTwo=invoiceFifteenListLoopTwo.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopTwo)
                                print('\n')   
                                invoiceFifteenListLoopThree=invoiceFifteenListLoopThree.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceFifteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceFifteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFifteenListLoopOne=invoiceFifteenListLoopOne.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopOne)
                                print('\n')
                                invoiceFifteenListLoopTwo=invoiceFifteenListLoopTwo.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopTwo)
                                print('\n')   
                                invoiceFifteenListLoopThree=invoiceFifteenListLoopThree.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopThree)
                                print('\n')
                                invoiceFifteenListLoopFour=invoiceFifteenListLoopFour.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceFifteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceFifteenListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceFifteenListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceFifteenListLoopOne=invoiceFifteenListLoopOne.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopOne)
                                print('\n')
                                invoiceFifteenListLoopTwo=invoiceFifteenListLoopTwo.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopTwo)
                                print('\n')   
                                invoiceFifteenListLoopThree=invoiceFifteenListLoopThree.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopThree)
                                print('\n')
                                invoiceFifteenListLoopFour=invoiceFifteenListLoopFour.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopFour)
                                print('\n')
                                invoiceFifteenListLoopFive=invoiceFifteenListLoopFive.flatten().tolist()
                                print( 'invoice15: ',invoiceFifteenListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceFifteenListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceFifteenListLoopFive, value_input_option='USER_ENTERED')

                        pass
                
        if os.path.exists(wegInvoice):
                with open(wegInvoice, 'r', encoding='utf-8') as f:
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0] ] # Primeiro valor da lista
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
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0] ] # Primeiro valor da lista
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
                        
                        if len(productCodeLoop) == 6:
    #Aqui nós retiramos os valores presentes da lista em Loop de dentro de uma lista e deixamos na lista principal
# a lista 'ref', por isso é usado o []+[]                                            
                                ref = [productCodeLoop[0][0], productCodeLoop[1][0], productCodeLoop[2][0], productCodeLoop[3][0], productCodeLoop[4][0], productCodeLoop[5][0] ] # Primeiro valor da lista
#PAREI AQUI NO DIA 08/08/2022. VALIDANDO MAIS DE UMA ORDEM DE COMPRA EM UMA NOTA DA WEG                                                          
                                procv = [ planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[0]), 'Transportadora'].iloc[0]]
                                procv1 = [ planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[1]), 'Transportadora'].iloc[0]]                              
                                procv2 = [ planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[2]), 'Transportadora'].iloc[0]]                              
                                procv3 = [ planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[3]), 'Transportadora'].iloc[0]]                              
                                procv4 = [ planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[4]), 'Transportadora'].iloc[0]]
                                procv5 = [ planilha01.loc[planilha01['Ref.'] == int(ref[5]), 'Código'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[5]), 'Nome do produto'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[5]), 'Prazo de Transporte(dias)'].iloc[0], planilha01.loc[planilha01['Ref.'] == int(ref[5]), 'Transportadora'].iloc[0]]

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
                                code5 = [procv5[0]]
                                product_name5 = [procv5[1]]
                                
                                
                                delivery_day = [procv[2]]
                                sum = (int(procv[2]))
                                delivery1 = delivery1 + sum 
                                pass

                                
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
                        invoiceList = np.array([nota+clienteOrderLoop[0]+date+['empty']+['empty']+['empty']+amountOrdered+['empty']+['empty']+['empty']])
                        invoiceList=invoiceList.flatten().tolist()
                        print('invoice: ', invoiceList)
                        print('\n')
                        
                        worksheet.append_row(invoiceList, value_input_option='USER_ENTERED')
                else:   
                        if len(clienteOrderLoop) == 1:
                                invoiceList = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])

                                invoiceList=invoiceList.flatten().tolist()
                                print( 'invoice: ',invoiceList)
                                print('\n')
                                        
                                worksheet.append_row(invoiceList, value_input_option='USER_ENTERED')
                        
                        elif len(clienteOrderLoop) == 2:
                                invoiceListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])

                                invoiceListLoopOne=invoiceListLoopOne.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopOne)
                                print('\n')
                                invoiceListLoopTwo=invoiceListLoopTwo.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopTwo)
                                print('\n')   
                                
                                worksheet.append_row(invoiceListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopTwo, value_input_option='USER_ENTERED')
                                
                        elif len(clienteOrderLoop) == 3:
                                invoiceListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])

                                invoiceListLoopOne=invoiceListLoopOne.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopOne)
                                print('\n')
                                invoiceListLoopTwo=invoiceListLoopTwo.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopTwo)
                                print('\n')   
                                invoiceListLoopThree=invoiceListLoopThree.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopThree)
                                print('\n')
                                worksheet.append_row(invoiceListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopThree, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 4:
                                invoiceListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceListLoopOne=invoiceListLoopOne.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopOne)
                                print('\n')
                                invoiceListLoopTwo=invoiceListLoopTwo.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopTwo)
                                print('\n')   
                                invoiceListLoopThree=invoiceListLoopThree.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopThree)
                                print('\n')
                                invoiceListLoopFour=invoiceListLoopFour.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopFour)
                                print('\n')
                                worksheet.append_row(invoiceListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopFour, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 5:
                                invoiceListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceListLoopOne=invoiceListLoopOne.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopOne)
                                print('\n')
                                invoiceListLoopTwo=invoiceListLoopTwo.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopTwo)
                                print('\n')   
                                invoiceListLoopThree=invoiceListLoopThree.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopThree)
                                print('\n')
                                invoiceListLoopFour=invoiceListLoopFour.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopFour)
                                print('\n')
                                invoiceListLoopFive=invoiceListLoopFive.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopFive)
                                print('\n')
                                worksheet.append_row(invoiceListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopFive, value_input_option='USER_ENTERED')

                        elif len(clienteOrderLoop) == 6:
                                invoiceListLoopOne = np.array([nota+clienteOrderLoop[0]+date+code+productCodeLoop[0]+product_name+amountOrderedLoop[0]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopTwo = np.array([nota+clienteOrderLoop[1]+date+code1+productCodeLoop[1]+product_name1+amountOrderedLoop[1]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopThree = np.array([nota+clienteOrderLoop[2]+date+code2+productCodeLoop[2]+product_name2+amountOrderedLoop[2]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopFour = np.array([nota+clienteOrderLoop[3]+date+code3+productCodeLoop[3]+product_name3+amountOrderedLoop[3]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopFive = np.array([nota+clienteOrderLoop[4]+date+code4+productCodeLoop[4]+product_name4+amountOrderedLoop[4]+delivery_day+final_delivery+shipping_company])
                                invoiceListLoopSix = np.array([nota+clienteOrderLoop[5]+date+code4+productCodeLoop[5]+product_name4+amountOrderedLoop[5]+delivery_day+final_delivery+shipping_company])

                                
                                invoiceListLoopOne=invoiceListLoopOne.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopOne)
                                print('\n')
                                invoiceListLoopTwo=invoiceListLoopTwo.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopTwo)
                                print('\n')   
                                invoiceListLoopThree=invoiceListLoopThree.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopThree)
                                print('\n')
                                invoiceListLoopFour=invoiceListLoopFour.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopFour)
                                print('\n')
                                invoiceListLoopFive=invoiceListLoopFive.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopFive)
                                print('\n')
                                invoiceListLoopSix=invoiceListLoopSix.flatten().tolist()
                                print( 'invoice: ',invoiceListLoopSix)
                                print('\n')
                                worksheet.append_row(invoiceListLoopOne, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopTwo, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopThree, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopFour, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopFive, value_input_option='USER_ENTERED')
                                worksheet.append_row(invoiceListLoopSix, value_input_option='USER_ENTERED')
        
                        pass
                        
                        
        else:
                print("Arquivo nao existe")
        weg = False
        break




#Usando a biblioteca OS é possível verificar se um arquivo xml existe e caso exista,
# o arquivo é excluído.


#While para apagar cada arquivo de XML
