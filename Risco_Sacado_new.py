from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta
import smtplib
from email.message import EmailMessage
from locale import setlocale, LC_ALL


class Risco:
    def __init__(self):
        self.cliente = 'utf-8'
        self.taxas = 0
        self.wb = load_workbook(filename='template_risco_sacado.xlsx')
        self.wb_cpgt = load_workbook(filename='template_cpgt_risco_sacado.xlsx')

    def interface(self):
        self.abrir_arq()
        active = True
        while active:
            self.cliente = input('Qual cliente você irá cadastrar? ').title()
            self.taxas = input('Qual a taxa? ')
            self.cpgt = input('Qual o prazo da condição de pagamento? ')
            self.banco = self.banco_1()
            self.lista_distr()
            self.abrir_plan_risco_sacado()
            self.abrir_plan_cpgt()
            self.salvar_arq()
            alerta = input('Prosseguir o cadastro ? \n(Pressione "enter" para continuar com os cadastros.\n'
                           'Caso deseje finalizar pressione "f" em seguida "enter".)-->')
            if alerta == 'f':
                active = False
        self.wb.close()
        self.wb_cpgt.close()
        self.enviar_email()

    def abrir_arq(self):
        return self.wb and self.wb_cpgt

    def salvar_arq(self):
        self.wb.save('risco_sacado(' + self.data_save_arquivo() + ').xlsx')
        self.wb_cpgt.save('Cadastro_em_lote(' + self.data_save_arquivo() + ').xlsx')

    def cpgt_terrestre(self):
        return 'ZD'+self.cpgt

    def cpgt_cabotagem(self):
        return 'ZC'+self.cpgt

    @staticmethod
    def banco_1():
        flag = True
        while flag:
            bancos = {'s': 'Santander', 'b': 'Bradesco', 'c': 'Citibank'}
            banco_marca = input('Escolha o banco?\n("s" para Santander, "b" para Bradesco e "c" para Citibank)'
                                ' + "enter" -->')
            if banco_marca in bancos:
                return bancos[banco_marca]
            else:
                print('Essa escolha não é possível, tente novamente!.')

    def lista_distr(self):
        distri = [self.cliente, self.taxas, self.cpgt_terrestre(), self.cpgt_cabotagem(), self.banco]
        return distri

    def distri_cliente_polo_produto(self):
        self.distribuioras = [Alesat,]
        Alesat = {8187:{1700:['PB.620','PB.6DH','PB.658']}
                  1740:{1400:['PB.620','PB.6DH','PB.658']}
                  }


    @staticmethod
    def centro_terrestre():
        terrestre = [1700, 1400, 1200, 1210, 1100, 1360, 1950, 1101, 1110, 1111, 1120, 1130]
        return terrestre

    @staticmethod
    def centro_cabotagem():
        cabotagem = [1401, 1211]
        return cabotagem

