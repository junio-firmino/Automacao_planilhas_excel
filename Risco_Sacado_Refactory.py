from openpyxl import load_workbook
import smtplib
from email.message import EmailMessage
from locale import setlocale, LC_ALL
import assists

# I will refactory This project, for do it I chose the design pattern Facade.

class Managerriscosacado:
    def __init__(self):
        print('Vamos iniciar o cadastro das condições do Risco Sacado para o Mês.')

    def create_plan_risco_sacado(self):
        self.plan_risco_sacado = Plan_risco_sacado()

    def create_plan_cpgt(self):
        self.plan_cpgt =


class Plan_risco_sacado:
    def plan_taxes(self):
        aba_act = self.wb.active
        self.lista_distr()
           for linha_plan in range(aba_act.max_row + 1, aba_act.max_row + 2):
            info = self.distri_cliente_polo_produto()[self.lista_distr()[0]]
            for fili, info_1 in info.items():
                for self.centro, prod in info_1.items():
                    for combust in prod:
                        aba_act.cell(row=linha_plan, column=1).value = fili  # Filial
                        aba_act.cell(row=linha_plan, column=2).value = self.cpgt_terrestre_cabotagem()  # CPGT
                        aba_act.cell(row=linha_plan, column=3).value = combust  # Produto
                        aba_act.cell(row=linha_plan, column=4).value = self.centro  # Centro
                        aba_act.cell(row=linha_plan, column=6).value = self.lista_distr()[1] + ' a.m.'  # Taxas
                        aba_act.cell(row=linha_plan, column=7).value = "%"
                        aba_act.cell(row=linha_plan, column=10).value = "A"
                        aba_act.cell(row=linha_plan, column=12).value = assists.data_inicio()  # Data inicial
                        aba_act.cell(row=linha_plan,
                                     column=13).value = assists.data_last_day_risco_sacado()  # Data final
                        aba_act.cell(row=linha_plan, column=14).value = self.lista_distr()[0]  # Cliente
                        aba_act.cell(row=linha_plan, column=15).value = self.encargos()  # Encargos
                        aba_act.cell(row=linha_plan, column=16).value = self.banco  # Banco
                        aba_act.cell(row=linha_plan, column=17).value = assists.data_cadastro()  # Data do cadastro
                        linha_plan += 1


class
