from PyQt5 import uic, QtWidgets
import sqlite3

def call_second_display():
    login_display.label_4.setText("")
    user_name = login_display.lineEdit.text()
    password = login_display.lineEdit_2.text()
    
    base = sqlite3.connect('E:/ALBACETE-AUTOMATION/DATABASE/database.db')
    cursor = base.cursor()
    try:
        cursor.execute("SELECT password FROM register WHERE login = '{}'".format(user_name) )
        password_db = cursor.fetchall()
        base.close()
    except:
        print("Erro ao validar as cedenciais")
        
    if  password ==  password_db[0][0]: #user_name == 'Lucas555'  and
        login_display.close()
        second_display.show()
    else:
        login_display.label_4.setText("Dados de login incorretos!")


      
def logout():
    second_display.close()
    login_display.show()
    
def open_register_display():
    register_display.show()
    
def register():
    name = register_display.lineEdit.text()
    login = register_display.lineEdit_2.text()
    password = register_display.lineEdit_3.text()
    c_password = register_display.lineEdit_4.text()
    
    if (password == c_password):
        try:
            base = sqlite3.connect('E:/ALBACETE-AUTOMATION/DATABASE/database.db')
            cursor = base.cursor()
            cursor.execute("CREATE TABLE IF NOT EXISTS register (name   text,login text,password text)")
            cursor.execute("INSERT INTO register VALUES ('"+name+"','"+login+"','"+password+"')")
            
            base.commit()   
            base.close()
            register_display.label.setText("Usuario cadastrado com sucesso")
            
        except sqlite3.Error as erro:
            print("Erro ao inserir os dados: ", erro)
    
    else:
        register_display.label.setText("As senhas digitadas estão diferentes")
        

app=QtWidgets.QApplication([])
login_display = uic.loadUi("E:/ALBACETE-AUTOMATION/FILES/Login/login_display.ui")
second_display = uic.loadUi("E:/ALBACETE-AUTOMATION/FILES/Login/second_display.ui")
register_display = uic.loadUi("E:/ALBACETE-AUTOMATION/FILES/Login/register_display.ui")

login_display.pushButton.clicked.connect(call_second_display)
second_display.pushButton_3.clicked.connect(logout)
login_display.lineEdit_2.setEchoMode(QtWidgets.QLineEdit.Password)
login_display.pushButton_2.clicked.connect(open_register_display)
register_display.pushButton.clicked.connect(register)

login_display.show()
app.exec()