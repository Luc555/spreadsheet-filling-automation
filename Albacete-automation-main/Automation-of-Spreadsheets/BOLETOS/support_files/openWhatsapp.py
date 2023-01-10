#import pyperclip
import time
import requests
import subprocess
import json
import requests
import pyperclip
import pyautogui
import calendar
from datetime import date

#Varíavel que armazena o dia atual
todayName = date.today() 
#Varíavel que armazena o nome do dia atual
todayNameShow = calendar.day_name[todayName.weekday()]
print(todayNameShow)

#Função de execução do Whatsapp
def openWhatsAppHome():
    
    # pyautogui.alert("O código vai começar. Não utilize nada do computador até o código finalizar!")
    #Tempo de espera do código
    time.sleep(2)
    #Abre o menu, como se o botão windows tivesse sido clicado
    pyautogui.press('winleft')
    #Tempo de espera do código
    time.sleep(3)
    pyautogui.PAUSE = 1.5
    #Escreve Whatsapp na busca já selecionada
    pyautogui.write('whatsapp')
    #Tempo de espera do código
    time.sleep(3)
    #Pressiona 'enter'
    pyautogui.press('enter')
    #Tempo de espera do código
    time.sleep(3)
    #Move o cursor do mouse para as coordenadas(x,y)
    pyautogui.moveTo(1180,50)
    #Clique
    pyautogui.click()
    #Tempo de espera do código
    time.sleep(3)
    pyautogui.write('@Dúvidas diversas')
    #Move o cursor do mouse para as coordenadas(x,y)
    pyautogui.moveTo(150,178)
    #Clique
    pyautogui.click()
    pyautogui.click()
    #Tempo de espera do código
    time.sleep(4)
    #Se o dia da semana estiver entre as posições abaixo, ele abre uma API
    if todayNameShow == 'Tuesday' or todayNameShow == 'Thursday':
        subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/support_files/translationApiValyrian.py", shell=True)
        
    #Caso contrário abre outro
    if todayNameShow == 'Monday' or todayNameShow == 'Wednesday' or todayNameShow == 'Friday':
        subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/support_files/translationApiDothrarki.py", shell=True)  
        
    #Pressiona 'enter'
    pyautogui.press('enter')
    #Move o cursor do mouse para as coordenadas(x,y)
    pyautogui.moveTo(1255,15)
    #Clique
    pyautogui.click()
    
    '''
    
    
    pyautogui.press('enter')
    pyautogui.moveTo(150,178)
    pyautogui.click()
    time.sleep(3)
    text_with_special_chars = 'Teste 01 - Garoto de programa\n 17/10/2022'
    pyperclip.copy(text_with_special_chars)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.moveTo(1000,15)
    pyautogui.click()


    #pyautogui.mouseDown()
    #pyautogui.moveTo(756,635)
'''
#Chama a função
openWhatsAppHome()