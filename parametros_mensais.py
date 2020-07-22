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
        self.ask = input("Qual tipo de contrato irá fazer o cadastro? ").title()
        if not isinstance(self.ask, LimiteAnswerPergunta):
            raise ValueError('Contrato não cadastrado.')

        return f'{self.ask}


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
        if self.answer_pergunta() in tabela.keys():
            return tabela.values()

    @staticmethod
    def data_inicio():
        return assists.data_inicio()

    def data_fim(self):
        pass