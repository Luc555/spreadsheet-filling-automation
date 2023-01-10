from PyQt5 import uic, QtWidgets
import subprocess
import os


def nf_emitidas():
    subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/nf_emitidas.py", shell=True)
    
def boletos():
    subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/duplicatas.py", shell=True)

def estoque_site():
    subprocess.call("C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/pedidos_abertos.py", shell=True)

app=QtWidgets.QApplication([])
menu_display = uic.loadUi("C:/Albacete-automation/Automation-of-Spreadsheets/MENU/Menu.ui")
menu_display.nfemitidas.clicked.connect(nf_emitidas)
menu_display.boletos.clicked.connect(boletos)
menu_display.estoque.clicked.connect(estoque_site)



menu_display.show()
app.exec()

