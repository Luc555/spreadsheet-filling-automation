#Binlioteca para tratamento de arquivos e diretórios
import os
#Validando erro de caminho não encontrado
#Para a criação da janela de erro
from time import sleep
from tkinter import *
from tkinter import messagebox


try:
        #Aviso do código que está rodando no momento
        print("Executando o código que apaga as os arquivos gerados na execução'")    
        #Para cada pasta, substasta ou arquivo presente dentro da pasta 'planilha_weg'
        #O código roda em recursão para cada uma das pastas, subpastas ou arquivos dentro da pasta acima
        for folder, subfolders, files in os.walk("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet"):
            #Mais um loop que para cada arquivo na lista de arquivos 
            for file in files:
                #Condição para arquivos terminados em .xls
                if file.endswith(".xls"):
                    #Cria uma variável que recebe o arquivo          
                    path = os.path.join(folder, file)  
                    #Exibe o arquivo selecionado
                    print('deleted : ', path ) 
                    #Apaga o arquivo
                    os.remove(path) 
                    
                #Condição para arquivos terminados em .xls    
                elif file.endswith(".xlsx"):
                    #Cria uma variável que recebe o arquivo        
                    path = os.path.join(folder, file)
                    #Exibe o arquivo selecionado  
                    print('deleted : ', path ) 
                    #Apaga o arquivo
                    os.remove(path)                              
            pass               

#Exceção para se o arquivo não for encontrado
except FileNotFoundError as error:
    #Variável recebe a mensagem padrão que será exibida no erro
    vmsg = "Não encontramos nem o arquivo e nem o diretório"
    # Variável recebe o erro
    tiposmg = error
    # Função para exibir mensagem
    def showMessage(tiposmg, msg):
        #Se a variável 'tiposmg' for igual ao erro
        if tiposmg == error:
            #Exibe a mensagem através de uma caixa com a mensagem padrão, definida mais acima
            messagebox.showerror(title="Sem notas fiscais", message=msg)
    #Chama a função
    showMessage(tiposmg, vmsg )