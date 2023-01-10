#Importando as bibliotecas necessárias para este trecho do código
import zipfile
import os
#Validando erro de caminho não encontrado
#Para a criação da janela de erro
from tkinter import *
from tkinter import messagebox
import shutil
import getpass
user = getpass. getuser().lower()


try:  
      print("Executando o código que extrai os arquivos do zip baixado do site")
      if os.path.exists("C:/Users/"+user+"/Downloads/wegInvoice.xml"):
            shutil.move("C:/Users/"+user+"/Downloads/wegInvoice.xml", "C:/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE")
            
            pass
      else:
            #Aqui nos criamos uma varíavel que localizará o arquivo de tal nome que será baixado
            #no diretório especificado.
            #Caminho computador Lucas casa 
            weg_zip = zipfile.ZipFile('C:/Users/'+user+'/Downloads/wegInvoices.zip')
            print(weg_zip)    
            
            #Caminho computador Lucas trabalho
            #weg_zip = zipfile.ZipFile('C:/Users/"+user+"/Downloads/wegInvoices.zip')

            #Aqui extraímos e colocamos na pasta tal 
            #Computador casa
            weg_zip.extractall('C:/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE')

            #Computador trabalho
            #weg_zip.extractall('C:/Albacete-automation/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE')
            #print('Deu certo!')
            weg_zip.close()
            
            #Computador casa
            """
            if os.path.exists("C:/Users/"+user+"/Downloads/wegInvoices.zip"):
                  os.remove("C:/Users/"+user+"/Downloads/wegInvoices.zip")
            """
                             
            
      
except FileNotFoundError as error:
      vmsg = "Não encontramos nem o arquivo e nem o diretório.\n Execução código Zip."
      tiposmg = error
      def showMessage(tiposmg, msg):
            if tiposmg == error:
                  messagebox.showerror(title="Sem notas fiscais", message=msg)
                  pass
      
      showMessage(tiposmg, vmsg )
      

