from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta
import smtplib
from email.message import EmailMessage
from locale import setlocale, LC_ALL
import assists


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
           # self.abrir_plan_cpgt()
            self.salvar_arq()
            alerta = input('Prosseguir o cadastro ? \n(Pressione "enter" para continuar com os cadastros.\n'
                           'Caso deseje finalizar pressione "f" em seguida "enter".)-->')
            if alerta == 'f':
                active = False
        self.wb.close()
        self.wb_cpgt.close()
        #self.enviar_email()

    def abrir_arq(self):
        return self.wb and self.wb_cpgt

    def salvar_arq(self):
        self.wb.save('risco_sacado).xlsx')
        self.wb_cpgt.save('Cadastro_em_lote.xlsx')

    def cpgt_terrestre_cabotagem(self):
        cpgt_centro = [1360, 1950]
        if cpgt_centro == 1360 or cpgt_centro == 1950:
            return 'ZC'+self.cpgt
        else:
            return 'ZD'+self.cpgt

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
        distri = [self.cliente, self.taxas, self.banco]
        return distri

    def clientes_distr(self):
        if self.lista_distr()[0] in self.distri_cliente_polo_produto():
            return self.lista_distr()[0]


    def distri_cliente_polo_produto(self):
        self.distribuidoras = {'Al': {8187: {1700: ['PB.620', 'PB.6DH', 'PB.658']},
                                      1740: {1400: ['PB.620', 'PB.6DH', 'PB.658']},
                                      4473: {1200: ['PB.620', 'PB.6DH', 'PB.658']},
                                      21699: {1100: ['PB.620', 'PB.6DH', 'PB.658']},
                                      4919: {1360: ['PB.6DH'], 1950: ['PB.6DH']},
                                      8429: {1101: ['PB.620', 'PB.6DH', 'PB.658']},
                                      1733: {1110: ['PB.620', 'PB.6DH', 'PB.658']},
                                      1732: {1111: ['PB.620', 'PB.6DH', 'PB.658']},
                                      1736: {1120: ['PB.620', 'PB.6DH', 'PB.658']},
                                      6833: {1130: ['PB.620', 'PB.6DH', 'PB.658']}},
                               'Cia': {455: {1400: ['PB.620', 'PB.6DH', 'PB.658']},
                                       18314: {1700: ['PB.620', 'PB.6DH', 'PB.658']},
                                       4150: {1100: ['PB.620', 'PB.6DH', 'PB.658']}},
                               'Ipp': {47: {1700: ['PB.620', 'PB.6DH', 'PB.658']},
                                       2093: {1400: ['PB.620', 'PB.6DH', 'PB.658']}},
                               'Mim': {17621: {1700: ['PB.620', 'PB.6DH', 'PB.658']}},
                               'Pet': {5142: {1360: ['PB.6DH'], 1950: ['PB.6DH']}},
                               'Rod': {7008: {1700: ['PB.620', 'PB.6DH', 'PB.658']},
                                       6815: {1400: ['PB.620', 'PB.6DH', 'PB.658']}},
                               'Rai': {49: {1700: ['PB.620', 'PB.6DH', 'PB.658']},
                                       2163: {1400: ['PB.620', 'PB.6DH', 'PB.658']},
                                       2153: {1200: ['PB.620', 'PB.6DH', 'PB.658'], 1210: ['PB.620', 'PB.6DH']},
                                       2150: {1100: ['PB.620', 'PB.6DH', 'PB.658']},
                                       2180: {1360: ['PB.6DH'], 1950: ['PB.6DH']},
                                       2155: {1101: ['PB.620', 'PB.6DH', 'PB.658']},
                                       18449: {1110: ['PB.620', 'PB.6DH', 'PB.658']},
                                       2186: {1111: ['PB.620', 'PB.6DH', 'PB.658']},
                                       2168: {1120: ['PB.620', 'PB.6DH', 'PB.658']},
                                       2157: {1130: ['PB.620', 'PB.6DH', 'PB.658']}}}

        #self.distribuidoras = [Cia, Al, Ipp, Mim, Pet, Rod, Rai]
        return self.distribuidoras

    @staticmethod
    def orgv():
        return "1001"

    @staticmethod
    def claros():
        return "02"

    @staticmethod
    def encargos():
        return '1,51% a.m'

    def abrir_plan_risco_sacado(self):
        aba_act = self.wb.active
        self.lista_distr()
        for linha_plan in range(2, 3):
            info = self.distri_cliente_polo_produto()[self.lista_distr()[0]]
            for fili, info_1 in info.items():
                for centro, prod in info_1.items():
                    for combust in prod:
                        aba_act.cell(row=linha_plan, column=2).value = self.lista_distr()[0]
                        aba_act.cell(row=linha_plan, column=3).value = self.orgv()
                        aba_act.cell(row=linha_plan, column=4).value = self.claros()
                        aba_act.cell(row=linha_plan, column=5).value = fili
                        aba_act.cell(row=linha_plan, column=6).value = centro
                        aba_act.cell(row=linha_plan, column=7).value = combust
                        aba_act.cell(row=linha_plan, column=8).value = self.cpgt_terrestre_cabotagem()
                        aba_act.cell(row=linha_plan, column=9).value = self.lista_distr()[1]
                        aba_act.cell(row=linha_plan, column=10).value = assists.data_inicio()
                        aba_act.cell(row=linha_plan, column=11).value = assists.data_last_day_risco_sacado()
                        aba_act.cell(row=linha_plan, column=12).value = self.lista_distr()[-1]
                        aba_act.cell(row=linha_plan, column=13).value = assists.data_cadastro()
                        aba_act.cell(row=linha_plan, column=14).value = self.encargos()
                        linha_plan += 1





x = Risco()
x.interface()