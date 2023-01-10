import shutil 
from tkinter import *
from tkinter import messagebox
import subprocess
import getpass

try:
    user = getpass. getuser().lower()  
    #Varíavel recebe o caminho do arquivo export na pasta Downloads
    source = "C:/Users/"+user+"/Downloads/export.xls"
    #Varíavel recebe o caminho de destino do arquivo export, na pasta planilha_weg
    destination = "C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/export.xls"
    #Move o arquivo do diretório de origem para o diretório de destino
    dest = shutil.move(source, destination)
    
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