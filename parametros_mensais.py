from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta
import smtplib
from email.message import EmailMessage
from locale import setlocale, LC_ALL


class Parametros:
    def __init__(self):
        self.montante = int

    def interface_client(self):
        self.abrir_arq()
        self.pergunta_1 = self.pergunta()
        self.montante = input('Qual o valor do parâmetro? ')
        self.list_trabalho()
        self.planilha_avulso()
        self.save_arq()

    def pergunta(self):
        contrato = ['Cgv', 'N4', 'Avulso']
        active = True
        while active:
            self.ask = input("Escolha o tipo de contrato? ").title()
            if self.ask in contrato:
                return self.ask
            else:
                print("não tem contrato")

    def list_trabalho(self):
        work = [self.pergunta_1, self.montante]
        return work

    def cliente_centro_produto(self):
        self.client = {
            15640: {'PB.620': [1100, 1101, 1150, 1160]},
            725: {'PB.620': [1100, 1101], 'PB.658': [1100, 1101, 1160], 'PB.6DH': [1160]},
            20347: {'PB.620': [1100], 'PB.658': [1100]},
            17644: {'PB.620': [1101], 'PB.658': [1101]},
            4168: {'PB.620': [1111], 'PB.658': [1111]},
            21254: {'PB.620': [1120]},
            16906: {'PB.620': [1120], 'PB.658': [1120], 'PB.6DH': [1120]},
            724: {'PB.620': [1120], 'PB.658': [1120], 'PB.6DH': [1120]},
            8425: {'PB.620': [1150]},
            15630: {'PB.620': [1150], 'PB.658': [1150], 'PB.6DH': [1150]},
            766: {'PB.620': [1160], 'PB.658': [1160], 'PB.6DH': [1160]},
            21933: {'PB.620': [1160]},
            8944: {'PB.620': [1160], 'PB.658': [1160], 'PB.6DH': [1160]},
            17984: {'PB.620': [1250], 'PB.658': [1250], 'PB.6DH': [1250]},
            16350: {'PB.620': [1360, 9044], 'PB.658': [1360, 9044], 'PB.6DH': [1360]},
            1123: {'PB.620': [1500, 1507], 'PB.658': [1500, 1507], 'PB.6DH': [1500, 1507], 'PB.650': [1500]},
            1125: {'PB.620': [1502, 9102], 'PB.658': [1502, 9102], 'PB.6DH': [1502, 9102], 'PB.650': [1502]},
            1124: {'PB.620': [1560, 9842, 9846, 9848], 'PB.658': [1560, 9842, 9846, 9848],
                   'PB.6DH': [1560, 9842, 9848]},
            10174: {'PB.620': [1560, 9848], 'PB.658': [1560, 9848], 'PB.6DH': [1560, 9848]},
            7169: {'PB.658': [1506]},
            6697: {'PB.658': [9842], 'PB.6DH': [9842]},
            4432: {'PB.650': [1050]},
            156: {'PB.650': [1400]},
            157: {'PB.650': [1400]},
            155: {'PB.650': [1423]}}
        return self.client

    def abrir_arq(self):
        self.wb = load_workbook(filename='template_PVA_PVS.xlsx')
        return self.wb

    def save_arq(self):
        self.wb.save('Carga' + self.list_trabalho()[0] + '.xlsx')

    @staticmethod
    def marca():
        return "x"

    @staticmethod
    def claros():
        return "02"

    @staticmethod
    def orgv():
        return "1001"

    def tipo_contrato(self):
        if self.list_trabalho()[0] == 'Cgv':
            return 'P'
        if self.list_trabalho()[0] == 'Avulso' or self.list_trabalho()[0] == 'N4':
            return 'N4'

    def grc4(self):
        self.condicoes_parametro = ['A', 'SP']
        return self.condicoes_parametro

    @staticmethod
    def material():
        produto = ['PB.620', 'PB.6DH', 'PB.658', 'PB.650']
        return produto

    @staticmethod
    def moeda():
        return 'BRL'

    @staticmethod
    def por():
        return "1"

    @staticmethod
    def unidade():
        return "M20"

    def tab(self):
        tabela = {'Cgv': 689, 'Avulsos': 525}
        for self.list_trabalho()[0], numero_tabela in tabela.items():
            return numero_tabela

    @staticmethod
    def data_inicial():
        data_inicio = dt.datetime.now() + relativedelta(day=1, months=1)
        data_inicio_return = data_inicio.strftime('%d.%m.%Y')
        return data_inicio_return

    @staticmethod
    def data_fim():
        data_last = dt.datetime.now() + relativedelta(day=31, months=1)
        data_return = data_last.strftime('%d.%m.%Y')
        return data_return

    def planilha_avulso(self):
        aba_avulso = self.wb.active
        self.list_trabalho()
        self.cliente_centro_produto()
        for linha_plan in range(3, len(self.grc4())+2):
            for condi in self.grc4():
                for filiais, carac in self.client.items():
                    for product, centre in carac.items():
                        for filial in centre:  # cada centro no dicionario
                            aba_avulso.cell(row=linha_plan, column=1).value = self.marca()
                            aba_avulso.cell(row=linha_plan, column=2).value = self.claros()
                            aba_avulso.cell(row=linha_plan, column=3).value = self.orgv()
                            aba_avulso.cell(row=linha_plan, column=6).value = self.tipo_contrato()
                            aba_avulso.cell(row=linha_plan, column=12).value = self.montante
                            aba_avulso.cell(row=linha_plan, column=13).value = self.moeda()
                            aba_avulso.cell(row=linha_plan, column=14).value = self.por()
                            aba_avulso.cell(row=linha_plan, column=15).value = self.unidade()
                            aba_avulso.cell(row=linha_plan, column=16).value = self.data_inicial()
                            aba_avulso.cell(row=linha_plan, column=17).value = self.data_fim()
                            aba_avulso.cell(row=linha_plan, column=18).value = self.tab()
                            aba_avulso.cell(row=linha_plan, column=4).value = condi
                            aba_avulso.cell(row=linha_plan, column=8).value = filiais
                            aba_avulso.cell(row=linha_plan, column=9).value = product
                            aba_avulso.cell(row=linha_plan, column=7).value = filial
                            linha_plan += 1


x = Parametros()
x.interface_client()




