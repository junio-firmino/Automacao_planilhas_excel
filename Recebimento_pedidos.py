import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver


class Abrir_Scd:
    def __init__(self,site):
        self.caminho = input("Qual o caminho do arquivo? ")
        self.brower = webdriver.Firefox(executable_path = self.caminho)
        self.brower.get(site)
        self.brower.implicitly_wait(5)

class Open_NVC(Abrir_Scd):
    def __init__(self,site):
        super.__init__(site)

class Close_NVC(Abrir_Scd):
    def __init__(self,site):
        super.__init__(site)

class Create_path:
    def path:
        pass

class Email:
    pass

class Choice_enginer:
    pass

class Enginer_open_NVC:
    pass

class Enginer_create_spreadsheet:
    pass

class Enginer_request_assent:
    pass