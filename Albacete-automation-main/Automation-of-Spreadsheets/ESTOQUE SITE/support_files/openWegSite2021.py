# Criar um código para avisar quando não tiver nota fiscal
# E validar a procura de notas na segunda, visto que tem notas que são emitidas na sexta-feira e no sábado
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from time import sleep
from datetime import datetime, timedelta
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
#Com a biblioteca 'Calendar' nós conseguimos exibior o dia da semana em que o código foi programado 
# para exibir, porém em Inglês
import calendar
from datetime import date
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from tkinter import *
from tkinter import messagebox
import time
import requests
import subprocess
import json
import requests
import pyperclip
import pyautogui


#Cria variável para trazer a data do dia anterior ao dia de execução do código
yesterdayName = date.today() - timedelta(1)
#Cria variável para trazer o dia de execução do código
todayName = date.today() 
#Cria variável para trazer o dia de execução do código e também o horário de execução
presentday = datetime.now() 
#Cria variável para trazer a data do dia anterior ao dia de execução do código
yesterday = presentday - timedelta(1)
#Cria variável para trazer a data de dois dias anteriores ao dia de execução do código 
twoDaysAgo = presentday - timedelta(2) 
#Cria variável para trazer a data de três dias anteriores ao dia de execução do código 
threeDaysAgo = presentday - timedelta(3)
#Cria variável para trazer a data do dia do amanhã ao dia de execução do código 
tomorrow = presentday + timedelta(1) 

#Mostrar o dia da semana do dia anterior, em inglês
yesterdayNameShow = calendar.day_name[yesterdayName.weekday()]
#Mostrar o dia da semana do dia atual, em inglês
todayNameShow = calendar.day_name[todayName.weekday()]

#Varíavel que converte a data recebida na variável threeDaysAgo(padrão estadunidense) para o 
# padrão de data brasileiro
tresDiasAtras = threeDaysAgo.strftime('%d/%m/%Y')
#Varíavel que converte a data recebida na variável twoDaysAgo(padrão estadunidense) para o 
# padrão de data brasileiro
doisDiasAtras = twoDaysAgo.strftime('%d/%m/%Y')
#Varíavel que converte a data recebida na variável yesterday(padrão estadunidense) para o 
# padrão de data brasileiro
ontem = yesterday.strftime('%d/%m/%Y')
#Varíavel que converte a data recebida na variável presentday(padrão estadunidense) para o 
# padrão de data brasileiro
hoje = presentday.strftime('%d/%m/%Y')

#Caminho do Google Chrome
options = webdriver.ChromeOptions()
options.binary_location = r"C:/Program Files/Google/Chrome Beta/Application/chrome.exe"

#Caminho do chromedriver
chrome_driver_binary = r"C:/Albacete-automation/Automation-of-Spreadsheets/chromedriver.exe"
driver = webdriver.Chrome(chrome_driver_binary, chrome_options=options)

#Aguarde 2 segundos
sleep(2)
#Abrindo o site através do webdriver instalado no diretório C
driver.get("https://www.weg.net/catalog/weg/BR/pt/login")

#Utilizando a tag Name para criar as variáveis abaixo. Temos as tags dos nomes dos elementos
user_path = 'j_username'
password_path ='j_password'

#Tempo de espera para o código ser carregado
sleep(2) 

#Criando outras variaveis para encontrar o elemento do html e retornar
user_element = driver.find_element(By.NAME, user_path)
password_element = driver.find_element(By.NAME, password_path)

#Usado para inserir nos elementos os respectivo dados que possibilitarão o login
user_element.send_keys("USUÁRIO DA EMPRESA")
password_element.send_keys("SENHA DA EMPRESA")

#Ativa o botão
button_element = WebDriverWait(driver, 20).until(
EC.element_to_be_clickable((By.XPATH, '//*[@id="loginForm"]/div[3]/button')))
button_element.click()
sleep(2) 

try:
        
    #Usado para carregar a página de notas fiscais 
    driver.get("https://www.weg.net/catalog/weg/BR/pt/research/open-orders")
    #Maximiza a janela do programa 
    driver.maximize_window()
    try:
        #Rola a barra lateal para a medida passada(horizontal, vertical)
        driver.execute_script('window.scrollBy(0, 350)')
    except:
        #Sem exceção
        None


    
    #A data que será inserida no campo de data inicial será a soma da data atual + 120
    #Essa soma é feito pela biblioteca datetime
    initialDate = '01/01/2021'
    finalDate = '31/12/2021'
 

    #Aqui capturamos o full xpath para a localização do campo no html
    initialDate_field = '/html/body/div[4]/div[1]/div/div[2]/div[7]/div/form/fieldset/div[2]/div/div/div[1]/input'
    initialDate_element = driver.find_element(By.XPATH, initialDate_field)
    #Aqui nos enviamos a data e a inserimos no campo
    initialDate_element.send_keys(initialDate)
    sleep(2)
    finalDate_field = '/html/body/div[4]/div[1]/div/div[2]/div[7]/div/form/fieldset/div[2]/div[1]/div/div[2]/input'
    finalDate_element = driver.find_element(By.XPATH, finalDate_field)
    #Aqui nos enviamos a data e a inserimos no campo
    finalDate_element.send_keys(finalDate)
    
    #Localização do Xpath do botão de busca
    searchButton = '/html/body/div[4]/div[1]/div/div[2]/div[7]/div/form/fieldset/div[3]/div/button'
    #searchButton_element = driver.find_element(By.XPATH, searchButton)
    #Aqui pedimos para que o código espere até o que o botão possa ser localizado
    searchButton_element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, searchButton)))
    searchButton_element.click()
    sleep(1)
    
    takeAllButton = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[7]/table/thead/tr/th[1]/input')))
    takeAllButton.click()
    sleep(1)
    
    xlsx_button = '/html/body/div[4]/div[1]/div/div[2]/div[7]/div[1]/div[2]/div[2]/form/button'
    #Aqui pedimos para que o código espere até o que o botão possa ser localizado
    xlsx_button_element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, xlsx_button)))
    xlsx_button_element.click()
    sleep(5)
    
            
        
#Tratamento de exceção            
except NoSuchElementException as error:
    #Criação da mensagem padrão
    vmsg = "A execução não pôde continuar, pois não teve emissão de nota fiscal no período."
    #Varíavel receb o erro
    tiposmg = error
    #Função para exibir a mensagem padrão
    def showMessage(tiposmg, msg):
        if tiposmg == error:
            #Caixa de mensagem criada
            messagebox.showerror(title="Sem notas fiscais", message=msg)
            subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/support_files/openWhatsappWhenThingsGetWrong.py", shell=True)
    #Chama a função
    showMessage(tiposmg, vmsg)
    #Encerra o driver
    driver.close()
        
        
