from selenium import webdriver
from time import sleep
from datetime import datetime, timedelta
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os.path
import numpy as np
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import *
from tkinter import messagebox
import subprocess



try:
     #Computador Compras ALbacete(Lucas)
      #exec(open("C:/Albacete-automation/Albacete-Automation/Automation-of-Spreadsheets/BOLETOS/support_files/openWebPage.py").read())
      #Código para abrir o site da Weg e fazer o download da planilha com as duplicatas
      subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/support_files/openWebPage.py", shell=True)
      
      #exec(open("C:/Albacete-automation/Albacete-Automation/Automation-of-Spreadsheets/BOLETOS/support_files/movingFiles.py").read())
      #Código para mover o 'export' da pasta Download do computador para a pasta planilha_weg
      subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/support_files/movingFiles.py", shell=True)
      
      #exec(open("C:/Albacete-automation/Albacete-Automation/Automation-of-Spreadsheets/BOLETOS/support_files/testDeletingRowsIfFindAnOrder.py").read())  
      #Código para transformação da planilha 'export' em dataframe, edição do dataframe, exclusão dos
      #dados da planilha 'Boleto - Planejamento-Motores'
      subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/support_files/editingSheet.py", shell=True)
      
      #Script que apaga todos os arquivos 'xlsx' e 'xls'
      subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/support_files/deletingSheets.py", shell=True)
      
      #Script que abre o Whatsapp e que consequentemente chama duas APIs, são assinalados por mensagem.
      subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/support_files/openWhatsapp.py", shell=True)

# Exceção de Aquivo não encontrado.        
except FileNotFoundError as error:
      #Mensagem padrão
      vmsg = "Não encontramos nem o arquivo e nem o diretório"
      # Váriavel vai receber a mensagem do erro, definida ao se passar a exceção(acima)
      tiposmg = error
      #Função para exibir mensagem, com os argumentos da mensagem e do erro
      def showMessage(tiposmg, msg):
            # Linha reduntante
            if tiposmg == error:
                  #Cria uma caixa de mensagem que irá exibir graficamente o erro e que o usuário deverá 
                  #clicar em 'Ok' para fechar  
                  messagebox.showerror(title="Sem arquivos ou diretório", message=msg)
      #Execução da função.
      showMessage(tiposmg, vmsg )

# Exceção do WebDriver(arquivo que permite) não encontrado.              
except WebDriverException as erro: 
      #Mensagem padrão      
      vmsg = "A versão do chromedriver está desatualizada.\n\n Favor atualizar a versão para que esta se adeque à versão do navegador:"
      # Váriavel vai receber a mensagem do erro, definida ao se passar a exceção(acima)
      tiposmg = erro
      #Função para exibir mensagem, com os argumentos da mensagem e do erro
      def showMessage(tiposmg, msg):
            #Transforma o erro em tipo String
            erroString = str(erro)
            #Variável receber uma String  vazia  
            delimiter = ''
            #Criar uma váriavel que junta todos os caracteres em uma única váriável
            erroStringList = delimiter.join(erroString)
            #Coletamos um grupo de caracteres(uma frase) da string anterior completa
            takingWhatMatters = erroStringList[118:128]
            #Linha redundante      
            if tiposmg == erro:
                  #Exibe uma caixa de mensagem  com a seguinte sequencia de termos
                  #mensagem é igual a mensagem da variável 'msg'(acima) e a frase coletada acima
                  messagebox.showerror(title="Chromedriver desatualizado", message=msg+takingWhatMatters+".\n\n Este é o link: https://chromedriver.chromium.org/downloads")
      #Executa a função              
      showMessage(tiposmg, vmsg)