import requests
import json
import pyperclip
import pyautogui


quote=[]
def quote_request():
    request = requests.get("https://api.gameofthronesquotes.xyz/v1/random/1")
    #request = requests.get("https://web-series-quotes-api.deta.dev")
    comments = json.loads(request.content)
    print(request.text)
    print(comments.keys())
    batata = comments.pop('sentence')
    quoteList = quote.append(batata)
    print(batata)
    
quote_request()
print(quote)
text = quote[0]


# importing the requests library
  
# defining the api-endpoint 
API_ENDPOINT = "https://api.funtranslations.com/translate/dothraki.json"
  
# your API key here
#No key
  
# your source code here
'''
print("Hello, world!")
a = 1
b = 2
print(a + b)
'''
  
# data to be sent to api
data = {"translated":'',
        "text": text,
        "translation": "dothraki"}
  
# sending post request and saving response as response object
r = requests.post(url = API_ENDPOINT, data = data)
comments = json.loads(r.content)
print(comments.keys())
batata = comments.pop('contents')
print(batata)
print(batata.keys())
tomate = batata.pop('translated')
print(tomate)
  
# extracting response text 
#pastebin_url = r.text
#print("The pastebin URL is:%s"%pastebin_url)
