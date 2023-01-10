from selenium import webdriver
#Pandas é uma biblioteca para manipulação de dados de planilhas e importamos criando um 
#apelido que é o pd
import pandas as pd
from tkinter import *
#Biblioteca para manipulação de dados de planilhas do GoogleSheets
import gspread
import subprocess
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import *
from tkinter import messagebox
from gspread.exceptions import SpreadsheetNotFound
from datetime import date
from time import sleep




try:
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
    worksheet = wks.worksheet("Estoque WEG (Site)")
    
    
    todayName = date.today()
    todayName = str(todayName)
    currentYear = todayName[0:4]
    
    def delitingData():
        

        #worksheet.delete_rows(2)
        #Linha que apaga as linhas da planilha a partir da terceira linha, pois nesta planilha
        #além do menu, a última linha restante não pode ser apagada. 
        #worksheet.row_count é a mesma coisa que o indice da planilha, ou seja, o total de linhas
        batata = worksheet.row_count
        print(batata)
        worksheet.delete_rows(3, worksheet.row_count)

        

    delitingData()
    
    #Apaga a linha repetida, que não havia sido deletada anteriormente    
        #worksheet.delete_rows(2)
        
    def dataFrame2021():
        planilha = pd.read_excel("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/export2021.xls", sheet_name="OpenOrders")
        # Get names of indexes for which column Stock has value No
        indexNames = planilha[ planilha['Setor de Atividade'] == 'Automação'].index
        indexNames1 = planilha[ planilha['Setor de Atividade'] == 'Construção Civil'].index
        # Delete these row indexes from dataFrame
        planilha.drop(indexNames, inplace=True)
        planilha.drop(indexNames1 , inplace=True)
        df = planilha.to_excel('C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/AtualizacaoExport.xlsx', index = False)
        df = pd.read_excel("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/AtualizacaoExport.xlsx", sheet_name="Sheet1")
        planilha = df.drop(['Setor de Atividade', 'Nome do Cliente', 'Data de Criação', 'Material do Cliente', 'Descrição do Produto', 'UM', 'Item', 'Item PO Cliente', 'Preço Unitário com Impostos', 'Valor Total com Impostos', 'Valor IPI (%)', 'Valor ST', 'Moeda'], axis=1)
        listPlannedBillingDate = planilha['Data do Faturamento Planejado'].tolist()
        
        
        listClientOrder = planilha['Pedido do Cliente'].tolist()
        listOrder = []
        for order in listClientOrder:
            order = str(order)
            if len(order) == 5:
                order = "0"+order
                int(order)
                listOrder.append(order)

        planilha['Pedido do Cliente'] = listOrder

        
        #Cria a lista para armazenar as quantidades
        numberList = []
        #Cria a lista para armazenar o termo 'em'
        emList = []
        #Cria a lista para armazenar a data de previsão
        dateList = []
        #Cria a lista para armazenar o mês de previsão
        monthList = []
        #Cria a lista para armazenar o ano de previsão
        yearList = []
        
        for item in listPlannedBillingDate:
            if len(item) == 15:
                number = item[:1]
                em = item[2:4]
                date = item[5:15]
                month = item[8:10]
                year = item[11:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
            if len(item) == 16:
                number = item[:2]
                em = item[3:5]
                date = item[6:16]
                month = item[9:11]
                year = item[12:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
            if len(item) == 17:
                number = item[:3]
                em = item[4:6]
                date = item[7:17]
                month = item[10:12]
                year = item[13:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
            
            if len(item) == 34:
                number1 = item[:1]
                int(number1)
                print(number)
                number2 = item[18:20]
                int(number2)
                number = number+number2
                str(number)
                print(number)
                em = item[2:4]
                date = item[5:15]
                month = item[8:10]
                year = item[11:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)

            if len(item) == 35:
                number = item[:3]
                print(number)
                em = item[4:6]
                date = item[7:17]
                month = item[10:12]
                year = item[13:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
        
        
        
        #Atualiza a coluna "Número" com a lista 'numberList'
        planilha['Número'] = numberList
        #Atualiza a coluna "Em" com a lista 'emList'
        planilha['Em'] = emList
        #Atualiza a coluna "Previsão" com a lista 'dateList'
        planilha['Previsão'] = dateList
        #Atualiza a coluna "Mês" com a lista 'monthList'
        planilha['Mês'] = monthList
        #Atualiza a coluna "Ano" com a lista 'yearList'
        planilha['Ano'] = yearList
        planilha01 = pd.read_excel("C:/Albacete-automation/DATABASE/Parametros-dos-motores.xlsx", sheet_name="Parâmetros dos Motores")

        
        listProductCode = planilha['Código do Produto'].tolist()
        j = 0
        codeList = []
        descriptionList = []
        
        for code in listProductCode:
            codeProcv = [ planilha01.loc[planilha01['Ref.'] == int(code), 'Código'].iloc[0]]
            codeProcv1 = codeProcv[0]
            stringCodeProcv = str(codeProcv1)
            lenStringCodeProcv = len(stringCodeProcv)
            if lenStringCodeProcv == 4:
                finalCode = '00'+stringCodeProcv
                int(finalCode)
                codeList.append(finalCode)
                    
            if lenStringCodeProcv == 5:
                finalCode = '0'+stringCodeProcv
                int(finalCode)
                codeList.append(finalCode)
            
            descriptionProcv = [ planilha01.loc[planilha01['Ref.'] == int(code), 'Nome do produto'].iloc[0]]
            descriptionProcv = descriptionProcv[0]
            description = str(descriptionProcv)
            descriptionList.append(description)
            j=j+1
            
        planilha['Código'] = codeList
        planilha['Descrição'] = descriptionList
            
        df = planilha.to_excel('C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/FinalSheet.xlsx', index = False)
        planilha = pd.read_excel("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/FinalSheet.xlsx", sheet_name="Sheet1")
        
        #Reorganiza o dataframe para que as colunas sejam iguais
        df = planilha[['Pedido do Cliente', 'Código', 'Código do Produto','Descrição','Quantidade', 'Quantidade em Processo Expedição', 'Quantidade em Aberto', 'Quantidade Disponível', 'Data prevista para Entrada no Estoque', 'Data do Faturamento Planejado', 'Número', 'Em', 'Previsão', 'Mês', 'Ano', 'Ordem de Venda WEG'    ]]
        df.fillna('', inplace=True) # atribui o novo valor ao df teres de lhe atribuir explicitamente o valor

        #df = planilha[['Pedido do Cliente', 'Código', 'Código do Produto','Descrição','Quantidade', 'Quantidade em Aberto', 'Data do Faturamento Planejado', 'Número', 'Em', 'Previsão', 'Mês', 'Ano', 'Ordem de Venda WEG'    ]]

        print(planilha)
        
        planilha = df.values.tolist()
        print(planilha)
        print(len(planilha))
        
        
        
        i=0
        #Para cada linha na lista df(lembrando que é uma lista composta de listas)
        for row in planilha:
            while i<len(planilha):
                #Insere na planilha 'BOLETO' a linha com a lista extraída de dentro da lista 'df' a medida em que 
                # soma o interador
                worksheet.append_row(planilha[i])
                #Soma mais um à variável interador
                i=i+1
            #Apaga a linha repetida, que não havia sido deletada anteriormente    
        worksheet.delete_rows(2)
        
        
    dataFrame2021()
    
    #Colocando para dormir 1 minutos, pois é o limite de requisições por minuto do Google
    sleep(60)
    
    def dataFrame():
        
        planilha = pd.read_excel("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/export"+currentYear+".xls", sheet_name="OpenOrders")
        # Get names of indexes for which column Stock has value No
        indexNames = planilha[ planilha['Setor de Atividade'] == 'Automação'].index
        indexNames1 = planilha[ planilha['Setor de Atividade'] == 'Construção Civil'].index
        # Delete these row indexes from dataFrame
        planilha.drop(indexNames, inplace=True)
        planilha.drop(indexNames1 , inplace=True)
        df = planilha.to_excel('C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/AtualizacaoExport.xlsx', index = False)
        df = pd.read_excel("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/AtualizacaoExport.xlsx", sheet_name="Sheet1")
        planilha = df.drop(['Setor de Atividade', 'Nome do Cliente', 'Data de Criação', 'Material do Cliente', 'Descrição do Produto', 'UM', 'Item', 'Item PO Cliente', 'Preço Unitário com Impostos', 'Valor Total com Impostos', 'Valor IPI (%)', 'Valor ST', 'Moeda'], axis=1)
        listPlannedBillingDate = planilha['Data do Faturamento Planejado'].tolist()
        
        
        listClientOrder = planilha['Pedido do Cliente'].tolist()
        listOrder = []
        for order in listClientOrder:
            order = str(order)
            if len(order) == 5:
                order = "0"+order
                int(order)
                listOrder.append(order)

        planilha['Pedido do Cliente'] = listOrder

        
        #Cria a lista para armazenar as quantidades
        numberList = []
        #Cria a lista para armazenar o termo 'em'
        emList = []
        #Cria a lista para armazenar a data de previsão
        dateList = []
        #Cria a lista para armazenar o mês de previsão
        monthList = []
        #Cria a lista para armazenar o ano de previsão
        yearList = []
        
        for item in listPlannedBillingDate:
            if len(item) == 15:
                number = item[:1]
                em = item[2:4]
                date = item[5:15]
                month = item[8:10]
                year = item[11:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
            if len(item) == 16:
                number = item[:2]
                em = item[3:5]
                date = item[6:16]
                month = item[9:11]
                year = item[12:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
            if len(item) == 17:
                number = item[:3]
                em = item[4:6]
                date = item[7:17]
                month = item[10:12]
                year = item[13:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
            
            if len(item) == 34:
                number1 = item[:1]
                int(number1)
                print(number)
                number2 = item[18:20]
                int(number2)
                number = number+number2
                str(number)
                print(number)
                em = item[2:4]
                date = item[5:15]
                month = item[8:10]
                year = item[11:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)

            if len(item) == 35:
                number = item[:3]
                print(number)
                em = item[4:6]
                date = item[7:17]
                month = item[10:12]
                year = item[13:]
                number.split()
                em.split()
                date.split()
                month.split()
                year.split()
                numberList.append(number)
                emList.append(em)
                dateList.append(date)
                monthList.append(month)
                yearList.append(year)
        
        
        
        #Atualiza a coluna "Número" com a lista 'numberList'
        planilha['Número'] = numberList
        #Atualiza a coluna "Em" com a lista 'emList'
        planilha['Em'] = emList
        #Atualiza a coluna "Previsão" com a lista 'dateList'
        planilha['Previsão'] = dateList
        #Atualiza a coluna "Mês" com a lista 'monthList'
        planilha['Mês'] = monthList
        #Atualiza a coluna "Ano" com a lista 'yearList'
        planilha['Ano'] = yearList
        planilha01 = pd.read_excel("C:/Albacete-automation/DATABASE/Parametros-dos-motores.xlsx", sheet_name="Parâmetros dos Motores")

        
        listProductCode = planilha['Código do Produto'].tolist()
        j = 0
        codeList = []
        descriptionList = []
        
        for code in listProductCode:
            codeProcv = [ planilha01.loc[planilha01['Ref.'] == int(code), 'Código'].iloc[0]]
            codeProcv1 = codeProcv[0]
            stringCodeProcv = str(codeProcv1)
            lenStringCodeProcv = len(stringCodeProcv)
            if lenStringCodeProcv == 4:
                finalCode = '00'+stringCodeProcv
                int(finalCode)
                codeList.append(finalCode)
                    
            if lenStringCodeProcv == 5:
                finalCode = '0'+stringCodeProcv
                int(finalCode)
                codeList.append(finalCode)
            
            descriptionProcv = [ planilha01.loc[planilha01['Ref.'] == int(code), 'Nome do produto'].iloc[0]]
            descriptionProcv = descriptionProcv[0]
            description = str(descriptionProcv)
            descriptionList.append(description)
            j=j+1
            
        planilha['Código'] = codeList
        planilha['Descrição'] = descriptionList
            
        df = planilha.to_excel('C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/FinalSheet.xlsx', index = False)
        planilha = pd.read_excel("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/FinalSheet.xlsx", sheet_name="Sheet1")
        
        #Reorganiza o dataframe para que as colunas sejam iguais
        df = planilha[['Pedido do Cliente', 'Código', 'Código do Produto','Descrição','Quantidade', 'Quantidade em Processo Expedição', 'Quantidade em Aberto', 'Quantidade Disponível', 'Data prevista para Entrada no Estoque', 'Data do Faturamento Planejado', 'Número', 'Em', 'Previsão', 'Mês', 'Ano', 'Ordem de Venda WEG'    ]]
        df.fillna('', inplace=True) # atribui o novo valor ao df teres de lhe atribuir explicitamente o valor

        #df = planilha[['Pedido do Cliente', 'Código', 'Código do Produto','Descrição','Quantidade', 'Quantidade em Aberto', 'Data do Faturamento Planejado', 'Número', 'Em', 'Previsão', 'Mês', 'Ano', 'Ordem de Venda WEG'    ]]

        print(planilha)
        
        planilha = df.values.tolist()
        print(planilha)
        print(len(planilha))
        
        
        
        i=0
        #Para cada linha na lista df(lembrando que é uma lista composta de listas)
        for row in planilha:
            while i<len(planilha):
                #Insere na planilha 'BOLETO' a linha com a lista extraída de dentro da lista 'df' a medida em que 
                # soma o interador
                worksheet.append_row(planilha[i])
                #Soma mais um à variável interador
                i=i+1
            #Apaga a linha repetida, que não havia sido deletada anteriormente    
        
        
    dataFrame()

except gspread.exceptions.SpreadsheetNotFound as e:
        #raise Exception("Trying to open non-existent or inaccessible spreadsheet document.")
        #Criação da mensagem padrão
        vmsg = "Tentando abrir um arquivo não existente ou inacessível."
        #Varíavel recebe o erro
        tiposmg = e
        #Função para exibir a mensagem padrão
        def showMessage(tiposmg, vmsg):
            if tiposmg == e:
                #Caixa de mensagem criada
                messagebox.showerror(title="Planilha não encontrada", message=vmsg)
        #Chama a função
        showMessage(tiposmg, vmsg)
        
except gspread.exceptions.WorksheetNotFound as e:
        #raise Exception("Trying to open non-existent sheet. Verify that the sheet name exists (%s)." % 'Estoque WEG (Site)')
        #Criação da mensagem padrão
        vmsg = "Tentando abrir uma planilha não existente. Verifique a planilha(aba) correta e altere a váriável worksheet no código ...ESTOQUE SITE/manipulatingDataFrame.py  "
        #Varíavel recebe o erro
        tiposmg = e
        #Função para exibir a mensagem padrão
        def showMessage(tiposmg, vmsg):
            if tiposmg == e:
                #Caixa de mensagem criada
                messagebox.showerror(title="Planilha não encontrada", message=vmsg)
        #Chama a função
        showMessage(tiposmg, vmsg)

except gspread.exceptions.APIError as e:
        if hasattr(e, 'response'):
            error_json = e.response.json()
            print(error_json)
            error_status = error_json.get("error", {}).get("status")
            vmsg = ""
            tiposmg = e
            
            if error_status == 'PERMISSION_DENIED':
                def showMessage(tiposmg, vmsg):
                    if tiposmg == e:
                        messagebox.showerror(title="Permissão negada", message="The Service Account does not have permission to read or write on the spreadsheet document. Have you shared the spreadsheet with %s?" % 'automatiza-o@automation-351815.iam.gserviceaccount.com automatiza-o@automation-351815.iam.gserviceaccount.com')        
                showMessage(tiposmg, vmsg )
                raise Exception("The Service Account does not have permission to read or write on the spreadsheet document. Have you shared the spreadsheet with %s?" % 'automatiza-o@automation-351815.iam.gserviceaccount.com automatiza-o@automation-351815.iam.gserviceaccount.com')
            if error_status == 'NOT_FOUND':
                def showMessage(tiposmg, vmsg):
                    if tiposmg == e:
                        messagebox.showerror(title="PLanilha não encontrada", message="Trying to open non-existent spreadsheet document. Verify the document id exists" + worksheet)        
                showMessage(tiposmg, vmsg )
                raise Exception("Tentando abrir uma planilha não existente. Verificar se a planilha correta e altere a váriável worksheet" )
            if error_status == 'INTERNAL':
                def showMessage(tiposmg, vmsg):
                    if tiposmg == e:
                        messagebox.showerror(title="Erro interno da API", message="Erro interno da API.\n Execute novamente.")        
                showMessage(tiposmg, vmsg )
                raise Exception("Erro interno da API. Execute o código novamente.")
            if error_status == 'INVALID_ARGUMENT':
                def showMessage(tiposmg, vmsg):
                    if tiposmg == e:
                        messagebox.showerror(title="Quantidade de linhas insuficientes", message="Tentando apagar linha de index 2(linha 3), mas só existem duas linhas.\n Favor adicionar mais uma linha e tentar novamente. \nAba Estoque Weg Site.")        
                showMessage(tiposmg, vmsg )
                raise Exception("Tentando apagar linha de index 2 (linha 3), mas só existem duas linhas. Favor adicionar mais uma linha e tentar novamente.\nAba Estoque Weg Site")
            raise Exception("The Google API returned an error: %s" % e)