import os
import os.path


try:
    os.makedirs("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/Validacao-nenhuma-emissão")
except OSError:
    #faz o que acha que deve se não for possível criar
    print( 'Error')
    
f = open("C:/Albacete-automation/Automation-of-Spreadsheets/NF EMITIDAS/Validacao-nenhuma-emissão/fileToStopProgram.txt", "x")

