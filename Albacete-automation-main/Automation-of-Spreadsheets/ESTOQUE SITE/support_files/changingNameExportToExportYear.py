import shutil
from datetime import date


todayName = date.today()
todayName = str(todayName)
currentYear = todayName[0:4]

shutil.move('C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/export.xls', 'C:/Albacete-automation/Automation-of-Spreadsheets/ESTOQUE SITE/weg_sheet/export'+currentYear+'.xls')