#import pyperclip
import time
import requests
import subprocess
import json
import requests
import pyperclip
import pyautogui

def openWhatsWhenTheresNoInvoice():
    
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
    text_with_special_chars = 'Wasted!! Não tivemos emissão de nota fiscal!!'
    programa = "Programa executado com sucesso"
    nf_emitida = "NF'Emitidas"
    pyperclip.copy(text_with_special_chars+'\n'+programa+'\n'+nf_emitida)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    pyautogui.moveTo(1255,15)
    pyautogui.click()

openWhatsWhenTheresNoInvoice()