import pyautogui
import time
import subprocess
subprocess.call('C:/Program Files (x86)/Alterdata/ERP/ShellERP.exe')
#pyautogui.alert("O código vai começar. Não utilize nada do computador até o código finalizar!")
pyautogui.PAUSE = 0.5


pyautogui.moveTo(567,38)
pyautogui.mouseDown()
pyautogui.moveTo(756,635)