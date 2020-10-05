import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
import os


class Abrir_Scd:
    def __init__(self,site,caminho):
        self.brower = webdriver.Firefox(executable_path = caminho)
        self.brower.get(site)
        self.brower.implicitly_wait(5)
        self.brower.maximize_window()
        self.login()

    def login(self):
        self.brower.find_element_by_css_selector("").send_keys('')
        self.brower.find_element_by_css_selector("").send_keys('')
        self.brower.find_element_by_name('').click()


class Open_NVC(Abrir_Scd):
    def __init__(self, site, caminho):
        super().__init__(site, caminho)
        #self.brower.close()


class Close_NVC(Abrir_Scd):
    def __init__(self,site):
        super().__init__(site)


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
    x=Open_NVC('http://www.anp.gov.br/', 'C:\\Users\\Jrfirmino Planejados\\Downloads\\geckodriver')