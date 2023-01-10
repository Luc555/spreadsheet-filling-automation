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
import calendar
from datetime import date
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from tkinter import *
from tkinter import messagebox
import os.path
import time
import requests
import subprocess
import json
import requests
import pyperclip
import pyautogui


yesterdayName = date.today() - timedelta(1)
todayName = date.today() 
presentday = datetime.now() 
yesterday = presentday - timedelta(1) 
twoDaysAgo = presentday - timedelta(2) 
threeDaysAgo = presentday - timedelta(3) 
tomorrow = presentday + timedelta(1) 

yesterdayNameShow = calendar.day_name[yesterdayName.weekday()]
todayNameShow = calendar.day_name[todayName.weekday()]


tresDiasAtras = threeDaysAgo.strftime('%d/%m/%Y')
doisDiasAtras = twoDaysAgo.strftime('%d/%m/%Y')
ontem = yesterday.strftime('%d/%m/%Y')
hoje = presentday.strftime('%d/%m/%Y')


options = webdriver.ChromeOptions()
options.binary_location = r"C:/Program Files/Google/Chrome Beta/Application/chrome.exe"
#Computador Trabalho
#chrome_driver_binary = r"C:/Albacete-automation/Albacete-automation/Automation-of-Spreadsheets/chromedriver.exe"

#Computador casa
chrome_driver_binary = r"C:/Albacete-automation/Automation-of-Spreadsheets/chromedriver.exe"
driver = webdriver.Chrome(chrome_driver_binary, chrome_options=options)



sleep(1)
#Abrindo o site através do webdriver instalado no diretório C
driver.get("https://www.weg.net/catalog/weg/BR/pt/login")

#Utilizando a tag Name para criar as variáveis
user_path = 'j_username'
password_path ='j_password'


#Tempo de espera para o código ser carregado
sleep(1) 

#Criando outras variaveis para encontrar o elemento do html e retornar
user_element = driver.find_element(By.NAME, user_path)
password_element = driver.find_element(By.NAME, password_path)

#Usado para inserir no elemento os respectivo dados que possibilitarão o login
user_element.send_keys("USUÁRIO DA ALBACETE")
password_element.send_keys("SENHA DA ALBACETE")

#Ativa o botão
button_element = WebDriverWait(driver, 20).until(
EC.element_to_be_clickable((By.XPATH, '//*[@id="loginForm"]/div[3]/button')))
button_element.click()
sleep(1) 

#Este if serve para a segunda-feira, pois temos de desconsiderar o dia anterior
# que é domingo e mesmo assim processar o sábado e a sexta-feira.
if todayNameShow == "Monday":

#Aqui temos o tratamento de exceção do tipo, não ter os elementos seguintes,
#não foram encontradas notas fiscais
    try:

        #Usado para carregar a página de notas fiscais 
        driver.get("https://www.weg.net/catalog/weg/BR/pt/research/invoices")
        driver.maximize_window()
        try:
            #driver.execute_script('window.scrollBy(0, 50)')
            driver.execute_script('window.scrollBy(0, 350)')
        except:
            None



        #Utilizando a biblioteca Datetime, aqui eu consigo pegar a data do dia, menos o dia anterior 
        initialDate = '//*[@id="initDate"]'
        finalDate = '//*[@id="finalDate"]'

        #Localiza-se onde os campos de datas se encontram na página específica do site 
        initialDate_element = driver.find_element(By.XPATH, '//*[@id="initDate"]')
        finalDate_element = driver.find_element(By.XPATH, '//*[@id="finalDate"]')

        sleep(1) 
        #Inserção no padrão dia/mês/ano
        initialDate_element.send_keys(threeDaysAgo.strftime('%d/%m/%Y'))
        finalDate_element.send_keys(presentday.strftime('%d/%m/%Y'))



        sleep(1)

        searchButton = '/html/body/div[4]/div[1]/div/div[2]/div[8]/div/form/fieldset/div[3]/div/a'
        #searchButton_element = driver.find_element(By.XPATH, searchButton)
        searchButton_element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, searchButton)))
        searchButton_element.click()
        sleep(2)

        select_element = driver.find_element(By.XPATH,"/html/body/div[4]/div[1]/div/div[2]/div[8]/table/thead/tr/th[1]/input")
        select_element.click()

        sleep(2)
        #buttonXml = driver.find_element(By.XPATH,"/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/button[1]")
        buttonXml = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/button[1]")))
        buttonXml.click()
        sucess = True
        sleep(1)
        vmsg = "A execução foi feita com sucesso.\n Os dados serão adicionados.\n Favor checar a planilha."
        tiposmg = sucess
        def showMessage(tiposmg, msg):
            if tiposmg == True:
                messagebox.showinfo(title="Tudo certo", message=msg)

        showMessage(True, vmsg )
        driver.close()
        
    except NoSuchElementException as error:
        print("\n")
        vmsg = "A execução não pôde continuar, pois não teve emissão de nota fiscal no período.\n Código de abertura do site"
        tiposmg = error
        def showMessage(tiposmg, msg):
            if tiposmg == error:
                messagebox.showerror(title="Sem notas fiscais", message=msg)
                

        showMessage(tiposmg, vmsg)
        
           
        driver.close()

else:    
        try:
        
            #Usado para carregar a página de notas fiscais 
            driver.get("https://www.weg.net/catalog/weg/BR/pt/research/invoices")
            driver.maximize_window()
            try:
                driver.execute_script('window.scrollBy(0, 350)')
            except:
                None



            #Utilizando a biblioteca Datetime, aqui eu consigo pegar a data do dia, menos o dia anterior 
            initialDate = '//*[@id="initDate"]'
            finalDate = '//*[@id="finalDate"]'

            #Localiza-se onde os campos de datas se encontram na página específica do site 
            initialDate_element = driver.find_element(By.XPATH, '//*[@id="initDate"]')
            finalDate_element = driver.find_element(By.XPATH, '//*[@id="finalDate"]')

            sleep(2) 
            #Inserção no padrão dia/mês/ano
            initialDate_element.send_keys(yesterday.strftime('%d/%m/%Y'))
            finalDate_element.send_keys(presentday.strftime('%d/%m/%Y'))

            
            searchButton = '/html/body/div[4]/div[1]/div/div[2]/div[8]/div/form/fieldset/div[3]/div/a'
            #searchButton_element = driver.find_element(By.XPATH, searchButton)
            searchButton_element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, searchButton)))
            searchButton_element.click()
            sleep(1)

            select_element = driver.find_element(By.XPATH,"/html/body/div[4]/div[1]/div/div[2]/div[8]/table/thead/tr/th[1]/input")
            select_element.click()

            sleep(2)
            #buttonXml = driver.find_element(By.XPATH,"/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/button[1]")
            buttonXml = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/button[1]")))
            buttonXml.click()
            sleep(5)
            
        except NoSuchElementException as error:
            print("\n")
            sleep(2)
            subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/support_files/creatingValidationToNoInvoiceStoppingProgram.py", shell=True)
            subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/support_files/openWhatsappWhenTheresNoInvoice.py", shell=True)
                    
            #exec(open("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/support_files/openWhatsapp.py").read())    
        
        except WebDriverException as erro:       
            vmsg = "A versão do chromedriver está desatualizada.\n\n Favor atualizar a versão para que esta se adeque à versão do navegador:"
            print(type(vmsg))
            tiposmg = erro
            def showMessage(tiposmg, msg):
                    batata = str(erro)
                    delimiter =  ''
                    batataList = delimiter.join(batata)
                    takingWhatMatters = batataList[118:128]
                    print(takingWhatMatters)
                        
                    if tiposmg == erro:
                        messagebox.showerror(title="Chromedriver desatualizado", message=msg+takingWhatMatters+".\n\n Este é o link: https://chromedriver.chromium.org/downloads")
                        print("Erro - exceção Webdriverexception: ", erro)
            showMessage(tiposmg, vmsg)    
            driver.close()
            
        
