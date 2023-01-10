#-*- coding:utf-8 -*-Voc  um pateta. Essa mensagem foi escrita atravs de um cdigo de programao pelo grande, lindo e maravilhosos, Lucas.
#import pyperclip
import time
import requests
import subprocess
import json
import requests
import pyperclip
import pyautogui

def openWhatsAppHome():
    
    # pyautogui.alert("O código vai começar. Não utilize nada do computador até o código finalizar!")
    time.sleep(5)
    pyautogui.press('winleft')
    time.sleep(5)
    pyautogui.PAUSE = 1.5
    pyautogui.write('whatsapp')
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.moveTo(1180,50)
    pyautogui.click()
    time.sleep(2)
    pyautogui.write('@Dúvidas diversas')
    pyautogui.moveTo(150,178)
    pyautogui.click()
    pyautogui.click()
    time.sleep(3)
    #text_with_special_chars = 'Teste 01 - Python\n 24/10/2022'
    exec(open("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/support_files/apiPython.py").read())
    pyautogui.press('enter')
    pyautogui.moveTo(1255,15)
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
openWhatsAppHome()