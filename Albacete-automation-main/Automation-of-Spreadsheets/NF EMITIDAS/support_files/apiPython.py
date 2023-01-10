import json
import requests
import pyperclip
import pyautogui




def buscar_dados():
    request = requests.get("https://api.adviceslip.com/advice")
    comments = json.loads(request.content)
    print(request.text)
    print(comments.keys())
    #print(comments['slip'])
    #Aqui nós extraimos o dicionário com as chaves 'id' e 'advice' de dentro do dicionário de chave
    #'slip'
    batata = comments.pop('slip')
    
    #Aqui printamos
    #print(batata)
    
    #Pegamos a lista de valores do dicionário e transformamos em uma lista
    listOfValues  = list(batata.values())
    
    #Aqui printamos a lista
    dailyPhrase = listOfValues[1]
    print(dailyPhrase)
    sucess = "Programa executado com sucesso!!"
    nf_emitida = "NF' Emitidas"
    pyperclip.copy(dailyPhrase+'\n'+sucess+'\n'+nf_emitida)
    pyautogui.hotkey('ctrl', 'v')
    
    #for val in listOfValues:
    #    print("Mensagem do dia: ",val)
        
    
buscar_dados()
