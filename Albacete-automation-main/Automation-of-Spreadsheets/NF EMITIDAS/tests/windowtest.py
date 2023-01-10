from tkinter import *

janela = Tk()
janela.title("Cotação Atual de Moedas")
texto = Label(janela, text="Clique no botão para ver as cotações de moedas")
texto.grid(column=0, row=0, padx=250, pady=250)

botao = Button(janela, text="Buscar cotações", ) #command=pegar_cotacoes)
botao.grid(column=0, row=1, padx=300, pady=300)

texto_resposta = Label(janela, text="")
texto_resposta.grid(column=0, row=2, padx=500, pady=500)

janela.mainloop()


class Application:
    def __init__(self, master=None):
        self.widget1 = Frame(master)
        self.widget1.pack(side=RIGHT)
        self.msg = Label(self.widget1, text="Primeiro widget")
        self.msg["font"] = ("Verdana", "10", "italic", "bold")
        self.msg.pack ()
        self.sair = Button(self.widget1)
        self.sair["text"] = "Sair"
        self.sair["font"] = ("Verdana", "10")
        self.sair["width"] = 10
        self.sair["command"] = self.widget1.quit
        self.sair.pack (side=RIGHT)
root = Tk()
Application(root)
root.mainloop()