from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta
import smtplib
from email.message import EmailMessage
from locale import setlocale, LC_ALL
import assists


class Parametros:
    def __init__(self):
        self.montante = int

    def interface_client(self):
        #self.abrir_arq()
        self.pergunta()
        self.montante = input('Qual o valor do parâmetro? ')
        self.planilha_avulso()

    def pergunta (self):
        active = True
        while active:
            self.ask = input("Escolha o tipo de contrato? ").title()
            tipo_contrato = ['Cgv', 'N4', 'Avulso']
            if self.ask in tipo_contrato:
                return self.ask
            else:
                print('Não é possível trabalhar com este contrato')

    def cliente_centro_produto(self):
        self.client = [15640, 725, 20347, 17644, 4168, 21254, 16906, 724, 8425, 15630, 766, 21933, 8944, 17984,
                       16350, 1123, 1125, 1124, 10174, 7169, 6697, 4432, 156, 157, 155]

        self.client[0] = {'PB.620': [1100, 1101, 1150, 1160]}        #15640
        self.client[1] = {'PB.620': [1100, 1101], 'PB.658': [1100, 1101, 1160], 'PB.6DH': [1160]}      #725
        self.client[2] = {'PB.620': [1100], 'PB.658': [1100]}        #20347
        self.client[3] = {'PB.620': [1101], 'PB.658': [1101]}        #17644
        self.client[4] = {'PB.620': [1111], 'PB.658': [1111]}        #4168
        self.client[5] = {'PB.620': [1120]}                        #21254
        self.client[6] = {'PB.620': [1120], 'PB.658': [1120], 'PB.6DH': [1120]}        #16906
        self.client[7] = {'PB.620': [1120], 'PB.658': [1120], 'PB.6DH': [1120]}        #724
        self.client[8] = {'PB.620': [1150]}                                        #8425
        self.client[9] = {'PB.620': [1150], 'PB.658': [1150], 'PB.6DH': [1150]}        #15630
        self.client[10] = {'PB.620': [1160], 'PB.658': [1160], 'PB.6DH': [1160]}       #766
        self.client[11] = {'PB.620': [1160]}                                       #21933
        self.client[12] = {'PB.620': [1160], 'PB.658': [1160], 'PB.6DH': [1160]}       #8944
        self.client[13] = {'PB.620': [1250], 'PB.658': [1250], 'PB.6DH': [1250]}       #17984
        self.client[14] = {'PB.620': [1360, 9044], 'PB.658': [1360, 9044], 'PB.6DH': [1360]}       #16350
        self.client[15] = {'PB.620': [1500, 1507], 'PB.658': [1500, 1507], 'PB.6DH': [1500, 1507], 'PB.650': [1500]}   #1123
        self.client[16] = {'PB.620': [1502, 9102], 'PB.658': [1502, 9102], 'PB.6DH': [1502, 9102], 'PB.650': [1502]}   #1125
        self.client[17] = {'PB.620': [1560, 9842, 9846, 9848], 'PB.658': [1560, 9842, 9846, 9848],
                           'PB.6DH': [1560, 9842, 9846, 9848]}       #1124
        self.client[18] = {'PB.620': [1560, 9848], 'PB.658': [1560, 9848], 'PB.6DH': [1560, 9848]}       #10174
        self.client[19] = {'PB.658': [1506]}       #7169
        self.client[20] = {'PB.658': [9842], 'PB.6DH': [9842]}       #6697
        self.client[21] = {'PB.650': [1050]}   #4432
        self.client[22] = {'PB.650': [1400]}   #156
        self.client[23] = {'PB.650': [1400]}   #157
        self.client[24] = {'PB.650': [1423]}   #155

        return self.client

    def abrir_arq(self):
        self.wb = load_workbook(filename='template_PVA_PVS.xlsx')

    def save_arq(self):
        self.wb.save('Carga'+self.pergunta()+'.xlsx')

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
        if self.pergunta() == 'Cgv':
            return 'P'
        if self.pergunta() == 'Avulso' or self.pergunta() == 'N4':
            return 'N4'

    @staticmethod
    def centro():
        centre = []
        return centre

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
        if self.pergunta() in tabela.keys():
            return tabela.values()

    @staticmethod
    def data_inicio():
        return assists.data_inicio()

    def data_fim(self):
        pass

    def planilha_avulso(self):
        #aba_avulso = self.wb.active
            for filiais in self.cliente_centro_produto():
                for product, centre in filiais.items():
                    print(product)  #
                    for filial in centre:
                        print(filial)   #



x = Parametros()
x.interface_client()
