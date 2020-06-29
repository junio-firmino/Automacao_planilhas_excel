from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta

class Risco():
    def __init__(self):
        self.cliente = True
        self.taxas = 0
        self.cpgt = True
        self.banco = True

    def interface(self):
        self.cliente = input('Qual cliente você irá cadastrar? ')
        self.taxas = input('Qual a taxa? ')
        self. cpgt = input('Qual condições de pagamento? ')
        self.banco = input('Qual banco escolhido? ')
        self.lista_distr()
        self.abrir_arq()


    def lista_distr(self):
        distri = []
        distri.append(self.cliente)
        distri.append(self.taxas)
        distri.append(self.cpgt)
        distri.append(self.banco)
        return distri

    def abrir_arq(self):
        wb = load_workbook(filename='teste_rs.xlsx')
        sheet_act = wb.active
        self.lista_distr()
        for linha_plan in range(2,sheet_act.max_row + 1):       # Tratamento na planilha das linhas
            empresa = sheet_act.cell(row= linha_plan, column= 1).value
            if empresa in self.lista_distr():
                sheet_act.cell(row = linha_plan, column= 3 ).value = self.lista_distr()[2]
                sheet_act.cell(row=linha_plan, column=4).value = self.lista_distr()[1]
                sheet_act.cell(row=linha_plan, column=5).value = self.data_inicio()
                sheet_act.cell(row=linha_plan, column=6).value = self.data_last_day()
                sheet_act.cell(row=linha_plan, column=7).value = self.lista_distr()[-1]
                sheet_act.cell(row=linha_plan, column=8).value = self.data_cadastro()

        wb.save('teste_rs.xlsx')

    #def abrir_arq_cpgt(self):

    def data_cadastro(self):
        data_cad = dt.datetime.now()
        return data_cad.strftime('%d.%m.%Y')

    def data_inicio(self):
        data_ini = dt.datetime.now() + relativedelta(months=1)

        data_1 = dt.datetime.now().strftime('01.%m.%Y')
        data_1_fort_date = dt.datetime.strptime(data_1,'%d.%m.%Y')

        data_2 = dt.datetime.now().strftime('21.%m.%Y')
        data_2_fort_date = dt.datetime.strptime(data_2,'%d.%m.%Y')

        data_3 = dt.datetime.now().strftime('%d.%m.%Y')
        data_3_fort_date = dt.datetime.strptime(data_3,'%d.%m.%Y')

        data_1_day = data_1_fort_date.day
        data_2_day = data_2_fort_date.day
        data_3_day = data_3_fort_date.day


        if data_3_day in range(0,(data_2_day - data_1_day)):
            return self.data_cadastro()
        else:
            return data_ini.strftime('01.%m.%Y')


    def data_last_day(self):
        last_day = dt.datetime.now() + relativedelta(day=31,months=2)
        return last_day.strftime('%d.%m.%Y')




x=Risco()
x.interface()
