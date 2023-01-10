import requests
import json
import pyperclip
import pyautogui

#Cria uma lista vazia
quote=[]
#Função para chamada de API
def quote_request():
    # Variável recebe a requisição de valor da API
    request = requests.get("https://poetrydb.org/random")
    #request = requests.get("https://web-series-quotes-api.deta.dev")
    #O método 'json.loads' é usado para converter JSON String para dicionário Python
    comments = json.loads(request.content)
    comments = comments[0]
    print(comments.keys())
    title = comments['author']
    author = comments['author']
    lines = comments['lines']
    linecount = comments['linecount']
    print('Título:',title)
    print('\n')
    str = '\n'.join(lines[:-1]) # gerar uma string com os items separados por virgula, com excecao do ultimo
    print(str)
    print('\n')
    print('Autor:',author)
    print('Número de linhas:',linecount)
    
    #Variável recebe frase
    sucess = "Programa executado com sucesso!!"
    #Variável recebe frase
    boleto = "'ESTOQUE WEG(SITE) '"
    #Utiliza biblioteca pyperclip para copiar o conteúdo passado entre parênteses
    pyperclip.copy(sucess+'\n'+boleto+'\n'+title+'\n'+str+'\n'+'\n'+author+'\n'+linecount)
    #Utiliza pyautogui para colar no campo selecionado naquele momento
    pyautogui.hotkey('ctrl', 'v')
        
#Chama a função    
quote_request()
