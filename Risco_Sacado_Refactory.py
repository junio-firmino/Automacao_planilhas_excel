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


class Planriscosacado:
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


class Plancpgt:
    def plan_cpgt(self):
        aba_act_cpgt = self.wb_cpgt.active
        self.lista_distr()
        for linha_cpgt in range(aba_act_cpgt.max_row + 1, aba_act_cpgt.max_row + 2):
            info = self.distri_cliente_polo_produto()[self.lista_distr()[0]]
            for fili, info_1 in info.items():
                for self.centro, prod in info_1.items():
                    for combust in prod:
                        aba_act_cpgt.cell(row=linha_cpgt, column=1).value = self.marca()
                        aba_act_cpgt.cell(row=linha_cpgt, column=2).value = self.claros()
                        aba_act_cpgt.cell(row=linha_cpgt, column=3).value = self.cpgt_terrestre_cabotagem()
                        aba_act_cpgt.cell(row=linha_cpgt, column=4).value = self.orgv()
                        aba_act_cpgt.cell(row=linha_cpgt, column=7).value = self.carencia_cpgt_terrestre_cabotagem()
                        aba_act_cpgt.cell(row=linha_cpgt, column=8).value = self.centro
                        aba_act_cpgt.cell(row=linha_cpgt, column=9).value = combust
                        aba_act_cpgt.cell(row=linha_cpgt, column=10).value = fili
                        aba_act_cpgt.cell(row=linha_cpgt, column=12).value = 1
                        aba_act_cpgt.cell(row=linha_cpgt, column=13).value = "BRL"
                        aba_act_cpgt.cell(row=linha_cpgt, column=14).value = 1
                        aba_act_cpgt.cell(row=linha_cpgt, column=15).value = "M20"
                        aba_act_cpgt.cell(row=linha_cpgt, column=16).value = "01.08.2020"
                        aba_act_cpgt.cell(row=linha_cpgt, column=17).value = "31.12.9999"
                        aba_act_cpgt.cell(row=linha_cpgt, column=18).value = self.tab()
                        aba_act_cpgt.cell(row=linha_cpgt, column=19).value = self.lista_distr()[0]
                        linha_cpgt += 1
