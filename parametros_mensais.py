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
        self.pergunta()
        self.montante = input('Qual o valor do parâmetro? ')

    def pergunta (self):
        active = True
        while active:
            self.ask = input("Escolha o tipo de contrato? ").title()
            tipo_contrato = ['Cg', 'N', 'Avul']
            if self.ask in tipo_contrato:
                return self.ask
            else:
                print('Não é possível trabalhar com este contrato')


    def cliente_centro_produto(self):
        self.filiais()[0] = {'PB.620': [1100, 1101, 1150, 1160]}        #15640
        self.filiais()[1] = {'PB.620': [1100, 1101], 'PB.658': [1100, 1101, 1160], 'PB,6DH': 1160}      #725
        self.filiais()[2] = {'PB.620': 1100, 'PB.658': 1100}        #20347
        self.filiais()[3] = {'PB.620': 1101, 'PB.658': 1101}        #17644
        self.filiais()[4] = {'PB.620': 1111, 'PB.658': 1111}        #4168
        self.filiais()[5] = {'PB.620': 1120}                        #21254
        self.filiais()[6] = {'PB.620': 1120, 'PB.658': 1120, 'PB.6DH': 1120}        #16906
        self.filiais()[7] = {'PB.620': 1120, 'PB.658': 1120, 'PB.6DH': 1120}        #724
        self.filiais()[8] = {'PB.620': 1150}                                        #8425
        self.filiais()[9] = {'PB.620': 1150, 'PB.658': 1150, 'PB.6DH': 1150}        #15630
        self.filiais()[10] = {'PB.620': 1160, 'PB.658': 1160, 'PB.6DH': 1160}       #766
        self.filiais()[11] = {'PB.620': 1160}                                       #21933
        self.filiais()[12] = {'PB.620': 1160, 'PB.658': 1160, 'PB.6DH': 1160}       #8944
        self.filiais()[13] = {'PB.620': 1250, 'PB.658': 1250, 'PB.6DH': 1250}       #17984
        self.filiais()[14] = {'PB.620': [1360, 9044], 'PB.658': [1360, 9044], 'PB.6DH': 1360}       #16350
        self.filiais()[15] = {'PB.620': [1500, 1507], 'PB.658': [1500, 1507], 'PB.6DH': [1500, 1507], 'PB.650': 1500}   #1123
        self.filiais()[16] = {'PB.620': [1502, 9102], 'PB.658': [1502, 9102], 'PB.6DH': [1502, 9102], 'PB.650': 1502}   #1125
        self.filiais()[17] = {'PB.620': [1560, 9842, 9846, 9848], 'PB.658': [1560, 9842, 9846, 9848],
                              'PB.6DH': [1560, 9842, 9846, 9848]}       #1124
        self.filiais()[18] = {'PB.620': [1560, 9848], 'PB.658': [1560, 9848], 'PB.6DH': [1560, 9848]}       #10174
        self.filiais()[19] = {'PB.658': 1506}       #7169
        self.filiais()[20] = {'PB.658': 9842, 'PB.6DH': 9842}       #6697
        self.filiais()[21] = {'PB.650': 1050}   #4432
        self.filiais()[22] = {'PB.650': 1400}   #156
        self.filiais()[23] = {'PB.650': 1400}   #157
        self.filiais()[24] = {'PB.650': 1423}   #155

    def filiais(self):
        client = [15640, 725, 20347, 17644, 4168, 21254, 16906, 724, 8425, 15630, 766, 21933, 8944, 17984,
                  16350, 1123, 1125, 1124, 10174, 7169, 6697, 4432, 156, 157, 155]
        return client

    def abrir_arq(self):
        self.wb = load_workbook(filename='template_PVA_PVS.xlsx')

    def save_arq(self):
        self.wb.save('Carga'+self.pergunta()+'.xlsx')

    @staticmethod
    def marca():
        return "x"

    @staticmethod
    def clear():
        return "02"

    @staticmethod
    def orgv():
        return "1001"

    def ty_contracty(self):
        if self.pergunta() == 'Cg':
            return 'P'
        if self.pergunta() == 'Avul' or self.pergunta() == 'N':
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
        tabela = {'Cg':689, 'Avul':525}
        if self.pergunta() in tabela.keys():
            return tabela.values()

    @staticmethod
    def data_inicio():
        return assists.data_inicio()

    def data_fim(self):
        pass


x = Parametros()
x.interface_client()
