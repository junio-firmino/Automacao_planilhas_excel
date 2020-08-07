from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta
import assists


class Parametros:
    def __init__(self):
        self.montante_cgv = 0
        self.montante_a = 0
        self.montante_sp = 0
        self.ask = str
        self.client = dict
        self.condicoes_parametro = list
        self.wb = load_workbook(filename='template_PVA_PVS.xlsx')

    def interface_client(self):
        self.abrir_arq()
        self.pergunta_1 = self.pergunta()
        self.montante()
        self.list_trabalho()
        self.planilha_referente_contrato()
        self.save_arq()

    def pergunta(self):
        contrato_escolhido = ['Cgv', 'N4', 'Avulso']
        active = True
        while active:
            self.ask = input("Escolha o tipo de contrato -->  ").title()
            if self.ask in contrato_escolhido:
                return self.ask
            else:
                print("Essa escolha não é possível, tente novamente!.")

    def planilha_referente_contrato(self):
        self.contrato = {'Cgv': self.planilha_cgv(), 'N4': self.planilha_n4(), 'Avulso': self.planilha_avulso()}
        return self.contrato[self.list_trabalho()[0]]

    def montante(self):
        if self.list_trabalho()[0] == 'Avulso' or self.list_trabalho()[0] == 'N4':
            self.montante_a = input('Qual o parâmetro para o Adicional (PVA)?')
            self.montante_sp = input('Qual o parâmetro para o Suplementar (PVS)?')
        else:
            self.montante_a = 0
            self.montante_sp = 0

        if self.list_trabalho()[0] == 'Cgv':
            self.montante_cgv = input('Qual o parâmetro CGV para o Adicional (PVA)? ')
        else:
            self.montante_cgv = 0

    def list_trabalho(self):
        work = [self.pergunta_1, self.montante_cgv, self.montante_a, self.montante_sp]
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
            155: {'PB.650': [1423]}
        }
        return self.client

    def produto_centro_cgv(self):
        self.produto_centro = {'PB.620': [1160, 1423, 1362, 1350, 1422, 1400, 1110, 1352, 1550, 1150, 1111, 1502,
                                          1365, 1560, 2540, 1100, 1050, 1360, 1101, 1421, 1700, 1250, 1120, 1070,
                                          1312, 1130, 9060, 1353, 1200, 1500, 1354, 1507, 1311, 1062, 1710],
                               'PB.650': [1500, 1400, 1423, 1050, 1560, 1200, 1550, 1502, 2540, 1365, 9060, 1070,
                                          1350, 1362, 1422],
                               'PB.658': [1160, 1423, 1362, 1350, 1422, 1400, 1110, 1352, 1550, 1150, 1111, 1502,
                                          1365, 1560, 2540, 1100, 1050, 1360, 1101, 1421, 1700, 1250, 1120, 1070,
                                          1312, 1130, 9060, 1353, 1200, 1500, 1354, 1507, 1311, 1062, 1710],
                               'PB.6DH': [1160, 1423, 1362, 1350, 1422, 1400, 1110, 1352, 1550, 1150, 1111, 1502,
                                          1365, 1560, 2540, 1100, 1050, 1360, 1101, 1421, 1700, 1250, 1120, 1070,
                                          1312, 1130, 9060, 1353, 1200, 1500, 1354, 1507, 1311, 1062, 1710],
        }

    def abrir_arq(self):
        return self.wb

    def save_arq(self):
        self.wb.save('PVA_PVS_' + self.list_trabalho()[0] + '_' + '(' + assists.data_cadastro() + ')' + '.xlsx')

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
        if self.list_trabalho()[0] == 'Avulso':
            self.condicoes_parametro = {'A': self.list_trabalho()[2], 'SP': self.list_trabalho()[3]}
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
        tabela = {'Cgv': 689, 'Avulso': 525}
        return tabela[self.list_trabalho()[0]]

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
        for linha_plan in range(3, len(self.grc4())+2):
            for condi, valor in self.grc4().items():
                for filiais, carac in self.cliente_centro_produto().items():
                    for product, centre in carac.items():
                        for numero_centre in centre:  # cada centro no dicionario
                            aba_avulso.cell(row=linha_plan, column=1).value = self.marca()
                            aba_avulso.cell(row=linha_plan, column=2).value = self.claros()
                            aba_avulso.cell(row=linha_plan, column=3).value = self.orgv()
                            aba_avulso.cell(row=linha_plan, column=6).value = self.tipo_contrato()
                            aba_avulso.cell(row=linha_plan, column=12).value = valor
                            aba_avulso.cell(row=linha_plan, column=13).value = self.moeda()
                            aba_avulso.cell(row=linha_plan, column=14).value = self.por()
                            aba_avulso.cell(row=linha_plan, column=15).value = self.unidade()
                            aba_avulso.cell(row=linha_plan, column=16).value = self.data_inicial()
                            aba_avulso.cell(row=linha_plan, column=17).value = self.data_fim()
                            aba_avulso.cell(row=linha_plan, column=18).value = self.tab()
                            aba_avulso.cell(row=linha_plan, column=4).value = condi
                            aba_avulso.cell(row=linha_plan, column=8).value = filiais
                            aba_avulso.cell(row=linha_plan, column=9).value = product
                            aba_avulso.cell(row=linha_plan, column=7).value = numero_centre
                            linha_plan += 1

    def planilha_n4(self):
        pass

    def planilha_cgv(self):
        pass


x = Parametros()
x.interface_client()
