import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver


class Abrir_Scd:
    caminho = input("Qual o caminho do arquivo? ")
    def __init__(self,site):
        self.brower = webdriver.Firefox(executable_path = caminho)
        self.brower.get(site)
        self.brower.implicitly_wait(5)

class Abrir_NVC(Abrir_Scd):
    def __init__(self,site):
        super.__init__(site)

class Fechar_NVC(Abrir_Scd):
    def __init__(self,site):
        super.__init__(site)

class Create_path:
    def path:
        pass

class Email:
    pass