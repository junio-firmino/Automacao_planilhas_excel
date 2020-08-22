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
        Al = {8187: {1700: ['PB.620','PB.6DH','PB.658']},
                  1740: {1400: ['PB.620','PB.6DH','PB.658']},
                  4473: {1200: ['PB.620','PB.6DH','PB.658']},
                  21699: {1100: ['PB.620','PB.6DH','PB.658']},
                  4919: {1360: ['PB.6DH'], 1950: ['PB.6DH']},
                  8429: {1101: ['PB.620','PB.6DH','PB.658']},
                  1733: {1110: ['PB.620','PB.6DH','PB.658']},
                  1732: {1111: ['PB.620','PB.6DH','PB.658']},
                  1736: {1120: ['PB.620','PB.6DH','PB.658']},
                  6833: {1130: ['PB.620','PB.6DH','PB.658']},}
        Cia = {455: {1400: ['PB.620','PB.6DH','PB.658']},
                    18314: {1700: ['PB.620','PB.6DH','PB.658']},
                    4150: {1100: ['PB.620','PB.6DH','PB.658']},}
        Ipp = {47: {1700: ['PB.620','PB.6DH','PB.658']},
               2093: {1400: ['PB.620','PB.6DH','PB.658']},}
        Mim = {17621: {1700: ['PB.620','PB.6DH','PB.658']}}
        Pet = {5142: {1360: ['PB.6DH'], 1950: ['PB.6DH']}}
        Rod = {7008: {1700: ['PB.620','PB.6DH','PB.658']},
                  6815: {1400: ['PB.620','PB.6DH','PB.658']}}
        Rai = {49: {1700: ['PB.620','PB.6DH','PB.658']},
                  2163: {1400: ['PB.620','PB.6DH','PB.658']},
                  2153: {1200: ['PB.620','PB.6DH','PB.658'], 1210: ['PB.620','PB.6DH']},
                  2150: {1100: ['PB.620','PB.6DH','PB.658']},
                  2180: {1360: ['PB.6DH'], 1950: ['PB.6DH']},
                  2155: {1101: ['PB.620','PB.6DH','PB.658']},
                  18449: {1110: ['PB.620','PB.6DH','PB.658']},
                  2186: {1111: ['PB.620','PB.6DH','PB.658']},
                  2168: {1120: ['PB.620','PB.6DH','PB.658']},
                  2157: {1130: ['PB.620','PB.6DH','PB.658']},}

        self.distribuidoras = [Al, Cia, Ipp, Mim, Pet, Rod,Rai]
        return self.distribuidoras


