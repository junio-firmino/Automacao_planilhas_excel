from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta
import smtplib
from email.message import EmailMessage
from locale import setlocale, LC_ALL
import assists
from enum import Enum

class LimiteAnswerPergunta(Enum):
    Cgv = 0
    N4 = 1
    Avulso = 2

class Parametros:
    def __init__(self):
        self.montante = int

    def interface_client(self):
        self.pergunta()
        self.montante = input('Qual o valor do parâmetro? ')

    def pergunta (self):
        self.ask = input("Escolha o tipo de contrato? ").title()
        if not isinstance(self.ask, LimiteAnswerPergunta):
            raise ValueError('Contrato não cadastrado.')

        return self.ask

    def cliente_centro_produto(self):
        self.filiais()[0] = {'PB.620': [1100, 1101, 1150, 1160]}
        self.filiais()[1] = {'PB.658': [1100, 1101, 1160],'PB.620': [1100, 1101], 'PB,6DH': 1160}
        self.filiais()[2] = {'PB.620': 1100, 'PB.658': 1100}
        self.filiais()[3] = {'PB.620': 1101, 'PB.658': 1101}
        self.filiais()[4] = {'PB.620': 1111, 'PB.658': 1111}
        self.filiais()[5] = {'PB.620': 1120}
        self.filiais()[6] = {'PB.620': 1120, 'PB.658': 1120, 'PB.6DH': 1120}
        self.filiais()[7] = {'PB.620': 1120, 'PB.658': 1120, 'PB.6DH': 1120}
        self.filiais()[8] = {'PB.620': 1150}
        self.filiais()[9] = {'PB.620': 1150, 'PB.658': 1150, 'PB.6DH': 1150}
        self.filiais()[10] = {'PB.620': 1160, 'PB.658': 1160, 'PB.6DH': 1160}
        self.filiais()[11] = {'PB.620': 1160}
        self.filiais()[12] = {'PB.620': 1160, 'PB.658': 1160, 'PB.6DH': 1160}
        self.filiais()[13] = {'PB.620': 1250, 'PB.658': 1250, 'PB.6DH': 1250}

    def filiais(self):
        client = [15640, 725, 20347, 17644, 4168, 21254, 16906, 724, 8425, 15630, 766, 21933, 8944, 17984,
                  16350, 1123, 1125, 1124, 10174, 7169, 6697, 4432, 156, 157, 155]
        return client

    def abrir_arq(self):
        pass

    def save_arq(self):
        pass

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
        pass

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
        tabela = {'Cgv':689, 'Avulsos':525}
        if self.pergunta() in tabela.keys():
            return tabela.values()

    @staticmethod
    def data_inicio():
        return assists.data_inicio()

    def data_fim(self):
        pass