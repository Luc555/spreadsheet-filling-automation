import requests
import json
import pyperclip
import pyautogui

#Criação de uma lista vazia
joke=[]
#Função chamada 'buscar_dados()'
def buscar_dados():
    # Variável recebe a requisição de valor da API
    request = requests.get("https://geek-jokes.sameerkumar.website/api?format=json")
    #O método 'json.loads' é usado para converter JSON String para dicionário Python
    comments = json.loads(request.content)
    #Exibe as chaves do dicionário
    print(comments.keys())
    
    #Insere a conteúdo presente dentro da chave 'joke' dentro da lista 'joke' criada no início do código
    joke.append(comments['joke'])
    

    """
    dailyPhrase = listOfValues[1]
    print(dailyPhrase)
    sucess = "Programa executado com sucesso!!"
    pyperclip.copy(dailyPhrase+'\n'+sucess)
    pyautogui.hotkey('ctrl', 'v')
    """
    
#Executa a função
buscar_dados()


#Exibe o valor presente dentro da lista
print(joke)