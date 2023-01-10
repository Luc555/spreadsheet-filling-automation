import os
#Validando erro de caminho não encontrado
#Para a criação da janela de erro
from time import sleep
from tkinter import *
from tkinter import messagebox
import getpass
user = getpass. getuser().lower()


try:
        sleep(5)
        print("Executando o código que apaga as os arquivos gerados na execução'")    
 
        for folder, subfolders, files in os.walk("C:/Albacete-automation/Automation-of-Spreadsheets/WEG-INVOICE"):
            for file in files:
                if file.endswith(".xml"):        
                    path = os.path.join(folder, file)  
                    print('deleted : ', path ) 
                    os.remove(path)                              
            pass
                
        if os.path.exists("C:/Users/"+user+"/Downloads/wegInvoices.zip"):
            os.remove("C:/Users/"+user+"/Downloads/wegInvoices.zip")
        pass

except FileNotFoundError as error:
    vmsg = "Não encontramos nem o arquivo e nem o diretório"
    tiposmg = error
    def showMessage(tiposmg, msg):
        if tiposmg == error:
            messagebox.showerror(title="Sem notas fiscais", message=msg)

    showMessage(tiposmg, vmsg )