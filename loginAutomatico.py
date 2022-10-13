import win32com.client
import sys
import subprocess
import time
from tkinter import *




class SapGui(object):
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(5)
        

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")       

        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection("Nome do sistema aqui", True) #nome do sistema / modulo
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize

    def SapLogin(self):

        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "###" #Nº mandante aqui
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "####" #Usuario aqui
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "####" #Senha aqui
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT" #Idioma
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(5)           

        except:
            print(sys.exc_info()[0])
        
  


if __name__ == '__main__':
    window = Tk()
    window.geometry("200x50")
    botao = Button(window, text="Login SAP", command= lambda: SapGui().SapLogin())
    botao.pack()
    mainloop()