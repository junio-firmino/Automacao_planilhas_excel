import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
import os


class Abrir_Scd:
    def __init__(self,site):
        self.caminho = input("Qual o caminho do arquivo? ")
        self.brower = webdriver.Firefox(executable_path = self.caminho)
        self.brower.get(site)
        self.brower.implicitly_wait(5)
        self.brower.maximize_window()

class Open_NVC(Abrir_Scd):
    def __init__(self,site):
        super.__init__(site)

class Close_NVC(Abrir_Scd):
    def __init__(self,site):
        super.__init__(site)

class Create_path:
    def path(self):
        pass

class Email:
    pass

class Choice_enginer:
    def __init__(self):
        self.escolha = input('Qual parte do processo de pedido você deseja executar?')

    def engineer(self):
        if self.escolha == '1':
            return Enginer_open_NVC()

        elif self.escolha == '2':
            return Enginer_create_spreadsheet()

        elif self.escolha == '3':
            return Enginer_request_assent()


class Enginer_open_NVC:
    pass

class Enginer_create_spreadsheet:
    pass

class Enginer_request_assent:
    pass

if __name__ == '__main__':
    x=Abrir_Scd("http://www.anp.gov.br/")