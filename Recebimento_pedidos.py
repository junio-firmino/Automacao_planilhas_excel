import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver


class Abrir_Scd:
    caminho = input("Qual o caminho do arquivo? ")
    def __init__(self,site):
        self.brower = webdriver.Firefox(executable_path=self.caminho)
        self.brower.get("http://www.anp.gov.br/")
        self.brower.implicitly_wait(5)

