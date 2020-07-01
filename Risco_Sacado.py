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
        wb = load_workbook(filename='template_risco_sacado.xlsx')
        sheet_act = wb.active
        self.lista_distr()
        for linha_plan in range(2,sheet_act.max_row + 1):       # Tratamento na planilha das linhas
            empresa = sheet_act.cell(row= linha_plan, column= 2).value
            if empresa in self.lista_distr():
                sheet_act.cell(row = linha_plan, column=8).value = self.lista_distr()[2]   # Preenche CPGT
                sheet_act.cell(row=linha_plan, column=9).value = self.lista_distr()[1]     # Prennche taxa
                sheet_act.cell(row=linha_plan, column=10).value = self.data_inicio()       # Prenche data inicio
                sheet_act.cell(row=linha_plan, column=11).value = self.data_last_day()     # Prenche data final
                sheet_act.cell(row=linha_plan, column=12).value = self.lista_distr()[-1]   # Prenche Banco
                sheet_act.cell(row=linha_plan, column=13).value = self.data_cadastro()     # Prenche data cadastro

        wb.save('risco_sacado_'+self.data_save()+'.xlsx')

    def abrir_arq_cpgt(self):
        wb_cpgt = load_workbook(filename= 'template_cpgt_risco_sacado.xlsx')
        aba_act = wb_cpgt.active
        self.lista_distr()
        for linha_cpgt in range(2,aba_act.max_row + 1):
            distribuidora = aba_act.cell(row=linha_cpgt,column=19).value
            if distribuidora in self.lista_distr():
                aba_act.cell(row=linha_cpgt,column=3).value = self.lista_distr()[2]
                aba_act.cell(row=linha_cpgt,column=16).value = self.data_inicio()
                aba_act.cell(row=linha_cpgt,column=17).value =



    def data_save(self):
        data_save_1 = dt.datetime.now()
        return data_save_1.strftime('%m_%y')

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
       data_last = dt.datetime.now() + relativedelta(day=31,months=1)
       data_last_1 = dt.datetime.now() + relativedelta(day=31)

       data_last_1_1 = dt.datetime.now().strftime('01.%m.%Y')
       data_last_1_1_fort_date = dt.datetime.strptime(data_last_1_1, '%d.%m.%Y')

       data_last_2 = dt.datetime.now().strftime('21.%m.%Y')
       data_last_2_fort_date = dt.datetime.strptime(data_last_2, '%d.%m.%Y')

       data_last_3 = dt.datetime.now().strftime('%d.%m.%Y')
       data_last_3_fort_date = dt.datetime.strptime(data_last_3, '%d.%m.%Y')

       data_last_1_day = data_last_1_1_fort_date.day
       data_last_2_day = data_last_2_fort_date.day
       data_last_3_day = data_last_3_fort_date.day

       if data_last_3_day in range(0, (data_last_2_day - data_last_1_day)):
           return data_last_1.strftime('%d.%m.%Y')
       else:
           return data_last.strftime('%d.%m.%Y')

    


x=Risco()
x.interface()
