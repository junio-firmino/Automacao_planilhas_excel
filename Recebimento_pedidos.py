import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
import os
from email.message import EmailMessage
from selenium.webdriver.common.keys import Keys


class Abrir_Scd:
    def __init__(self, site, caminho):
        self.brower = webdriver.Ie(executable_path=caminho)
        self.brower.get(site)
        self.brower.maximize_window()


class Login(Abrir_Scd):
    def __init__(self, site, caminho):
        super().__init__(site, caminho)
        self.brower.find_element_by_name("txt_user_login").send_keys('')
        self.brower.find_element_by_name("pwd_user_password").send_keys('')
        self.brower.find_element_by_name("//button[text()='Entar']").click()

        # x = self.brower.find_element_by_name("pwd_user_password")
        # x.send_keys(Keys.TAB)
        # x.submit()


class Close_NVC(Abrir_Scd):
    def __init__(self, site):
        super().__init__(site)
        # self.brower.close()


class Create_path:
    def path(self):
        pass


class Email:
    def __init__(self, email):
        self.email = email

    def config(self):
        smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
        smtpobj.starttls()
        fro = ''
        to = ''

        smtpobj.login(fro, '')
        msg = EmailMessage()
        msg['From'] = fro
        msg['To'] = to
        msg['Subject'] = f'Recebimento de Pedido mês {data_visivel}.'

        msg.set_content(
            f'Prezados\n\nSegue abaixo a planilha com as taxas dos clientes que utilizarão'
            f' as condições de pagamento na modalidade risco sacado para o mês de {data_visivel}.')
        paths = ['risco_sacado(' + assists.data_cadastro() + ').xlsx',
                 'Cadastro_em_lote_RS(' + assists.data_cadastro() + ').xlsx']
        for path in paths:
            caminho = open(path, 'rb')
            arq_data = caminho.read()
            arq_name = caminho.name
            msg.add_attachment(arq_data, maintype='application', subtype='octet-stream', filename=arq_name)

        smtpobj.send_message(msg)
        smtpobj.quit()
        print('Email enviado!!!!')


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


class Enginer_open_NVC(Abrir_Scd):
    def __init__(self, site, caminho):
        super().__init__(site, caminho)
        Login(site, caminho)
        self.brower.implicitly_wait(10)


class Enginer_create_spreadsheet:
    pass


class Enginer_request_assent:
    pass


if __name__ == '__main__':
    x = Open_NVC(
        'https://servicoca.petrobras.com.br/fwca/pages/AuthenticationForm.jsp?successfulUrl=https://sgcd.petrobras.com.br:443/ASPX/Home/frmHome.aspx&ssoEnabled=False&applicationCatalogId=S223&appEnvUid=493299&integratedAuthenticationEnabled=False&logonPage=',
        'C:\\Users\\e43k\\Downloads\\IEDriverServer.exe')
    x.operacion_NVC()
