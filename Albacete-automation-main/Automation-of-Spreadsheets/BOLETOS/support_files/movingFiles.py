import shutil
import getpass
user = getpass. getuser().lower()
#Varíavel recebe o caminho do arquivo export na pasta Downloads
source = "C:/Users/"+user+"/Downloads/export.xls"
print(source)
#Varíavel recebe o caminho de destino do arquivo export, na pasta planilha_weg
destination = "C:/Albacete-automation/Automation-of-Spreadsheets/BOLETOS/planilha_weg/export.xls"
#Move o arquivo do diretório de origem para o diretório de destino
dest = shutil.move(source, destination)
