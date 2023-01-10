import requests
import json
import pyperclip
import pyautogui

#Cria uma lista vazia
quote=[]
#Função para chamada de API
def quote_request():
    # Variável recebe a requisição de valor da API
    request = requests.get("https://api.gameofthronesquotes.xyz/v1/random/1")
    #request = requests.get("https://web-series-quotes-api.deta.dev")
    #O método 'json.loads' é usado para converter JSON String para dicionário Python
    comments = json.loads(request.content)
    #Exibe as chaves do dicionário
    print(comments.keys())
    #Extrai o conteudo da chave 'sentence'
    sentence = comments.pop('sentence')
    #Envia para a lista
    quote.append(sentence)

#Chama a função    
quote_request()
#Extrair o texto de dentro da lista
text = quote[0]
  
# Definindo a api-endpoint 
API_ENDPOINT = "https://api.funtranslations.com/translate/dothraki.json"
  
#Essa API não utiliza chave, porém tem limite de 5 requisições por hora 
  
# Dados enviados para a API
data = {"translated":'',
        "text": text,
        "translation": "dothraki"}
  
# Enviando post request e salvando a resposta como dicionário Python
r = requests.post(url = API_ENDPOINT, data = data)
comments = json.loads(r.content)

#Chaves do dicionário
chaves = comments.keys()
#Variável recebe a variável 'chaves'
s =chaves;
#Transforma em lista
s= list(s)
#Cria uma lista vazia
translatedList = []

#Se o valor retirado da lista for igual a condição logo abaixo, o código irá retirar o conteudo de dentro
#daquela chave e irá printar a mensagem
if s[0] == 'error':
        error = comments.pop('error')
        print(error)
        print(error.keys())
        message = error.pop('message')
        print(message)
   
#Se o valor retirado da lista for igual a condição logo abaixo, o código irá retirar o conteudo de dentro
#daquela chave e irá printar a mensagem e depois irá inserir o conteúdo dentro da lista   
if s[0] == 'success':        
        
        print('\n')
        print(comments.keys())
        contents = comments.pop('contents')
        print(contents)
        translated = contents.pop('translated')
        print(translated)
        translatedList.append(translated)
        pass

#Retira a frase contida dentro da lista
dailyPhrase = translatedList[0]
#Variável recebe frase
sucess = "Programa executado com sucesso!!"
#Variável recebe frase
boleto = "'Planilha Boletos'"
#Utiliza biblioteca pyperclip para copiar o conteúdo passado entre parênteses
pyperclip.copy(text+'\n'+'Dothraki: '+dailyPhrase+'\n'+sucess+'\n'+boleto)
#Utiliza pyautogui para colar no campo selecionado naquele momento
pyautogui.hotkey('ctrl', 'v')




 
 
