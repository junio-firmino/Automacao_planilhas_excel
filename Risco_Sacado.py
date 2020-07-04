from openpyxl import load_workbook
import datetime as dt
from dateutil.relativedelta import relativedelta
import smtplib
from email.message import EmailMessage


class Risco:
    def __init__(self):
        self.cliente = 'utf-8'
        self.taxas = 0
        self.cpgt = 'utf-8'

    def interface(self):
        self.abrir_arq()
        active = True
        while active:
            self.cliente = input('Qual cliente você irá cadastrar? ').title()
            self.taxas = input('Qual a taxa? ')
            self.cpgt = input('Qual condições de pagamento? ').upper()
            self.banco = self.banco_1()
            self.lista_distr()
            self.abrir_plan()
            self.abrir_plan_cpgt()
            alerta = input('Prosseguir o cadastro ?')
            if alerta == 'n':
                active = False
        self.enviar_email()

    def abrir_arq(self):
        self.wb = load_workbook(filename='template_risco_sacado.xlsx')
        self.wb_cpgt = load_workbook(filename='template_cpgt_risco_sacado.xlsx')

    def banco_1(self):
        banco_marca = input('Qual banco escolhido? ')
        if banco_marca == 's':
            return 'Santander'
        if banco_marca == 'b':
            return 'Bradesco'

    def lista_distr(self):
        distri = [self.cliente, self.taxas, self.cpgt, self.banco]
        return distri

    def abrir_plan(self):
        sheet_act = self.wb.active
        self.lista_distr()
        for linha_plan in range(2, sheet_act.max_row + 1):  # Tratamento na planilha das linhas
            empresa = sheet_act.cell(row=linha_plan, column=2).value
            if empresa in self.lista_distr():
                sheet_act.cell(row=linha_plan, column=8).value = self.lista_distr()[2]  # Preenche CPGT
                sheet_act.cell(row=linha_plan, column=9).value = self.lista_distr()[1]  # Preenche taxa
                sheet_act.cell(row=linha_plan, column=10).value = self.data_inicio()  # Preenche data inicio
                sheet_act.cell(row=linha_plan, column=11).value = self.data_last_day()  # Preenche data final
                sheet_act.cell(row=linha_plan, column=12).value = self.lista_distr()[-1]  # Preenche Banco
                sheet_act.cell(row=linha_plan, column=13).value = self.data_cadastro()  # Preenche data cadastro

        self.wb.save('risco_sacado(' + self.data_save() + ').xlsx')

    def abrir_plan_cpgt(self):
        aba_act = self.wb_cpgt.active
        self.lista_distr()
        for linha_cpgt in range(2, aba_act.max_row + 1):
            distribuidora = aba_act.cell(row=linha_cpgt, column=19).value
            if distribuidora in self.lista_distr():
                aba_act.cell(row=linha_cpgt, column=3).value = self.lista_distr()[2]
                aba_act.cell(row=linha_cpgt, column=16).value = self.data_inicio()
                aba_act.cell(row=linha_cpgt, column=17).value = self.data_last_day_cpgt()
                aba_act.cell(row=linha_cpgt, column=7).value = self.carencia_cpgt()

        self.wb_cpgt.save('Cadastro_CPGT_RS(' + self.data_save() + ').xlsx')

    def carencia_cpgt(self):
        condicoes_cpgt = self.lista_distr()[2]
        list_separador = condicoes_cpgt.split('D')
        valor_separado = list_separador[1]
        resultado = int(valor_separado) - 1
        return resultado

    def data_save(self):
        data_save_1 = dt.datetime.now()
        return data_save_1.strftime('%d_%m_%y')

    def data_cadastro(self):
        data_cad = dt.datetime.now()
        return data_cad.strftime('%d.%m.%Y')

    def data_inicio(self):
        data_ini = dt.datetime.now() + relativedelta(months=1)
        data_1 = dt.datetime.now().strftime('01.%m.%Y')
        data_1_transf_date = dt.datetime.strptime(data_1, '%d.%m.%Y')
        data_2 = dt.datetime.now().strftime('21.%m.%Y')
        data_2_transf_date = dt.datetime.strptime(data_2, '%d.%m.%Y')
        data_3 = dt.datetime.now().strftime('%d.%m.%Y')
        data_3_transf_date = dt.datetime.strptime(data_3, '%d.%m.%Y')
        data_1_day = data_1_transf_date.day
        data_2_day = data_2_transf_date.day
        data_3_day = data_3_transf_date.day
        if data_3_day in range(0, (data_2_day - data_1_day)):
            return self.data_cadastro()
        else:
            return data_ini.strftime('01.%m.%Y')

    def data_last_day(self):
        data_last = dt.datetime.now() + relativedelta(day=31, months=1)
        data_last_1 = dt.datetime.now() + relativedelta(day=31)
        data_last_1_1 = dt.datetime.now().strftime('01.%m.%Y')
        data_last_1_1_transf_date = dt.datetime.strptime(data_last_1_1, '%d.%m.%Y')
        data_last_2 = dt.datetime.now().strftime('21.%m.%Y')
        data_last_2_transf_date = dt.datetime.strptime(data_last_2, '%d.%m.%Y')
        data_last_3 = dt.datetime.now().strftime('%d.%m.%Y')
        data_last_3_transf_date = dt.datetime.strptime(data_last_3, '%d.%m.%Y')
        data_last_1_day = data_last_1_1_transf_date.day
        data_last_2_day = data_last_2_transf_date.day
        data_last_3_day = data_last_3_transf_date.day
        if data_last_3_day in range(0, (data_last_2_day - data_last_1_day)):
            return data_last_1.strftime('%d.%m.%Y')
        else:
            return data_last.strftime('%d.%m.%Y')

    def data_last_day_cpgt(self):
        data_last_cpgt = dt.datetime.now() + relativedelta(day=1, months=3)
        data_last_1_cpgt = dt.datetime.now() + relativedelta(day=1, months=2)
        data_last_1_1_cpgt = dt.datetime.now().strftime('01.%m.%Y')
        data_last_1_1_cpgt_transf_date = dt.datetime.strptime(data_last_1_1_cpgt, '%d.%m.%Y')
        data_last_2_cpgt = dt.datetime.now().strftime('21.%m.%Y')
        data_last_2_cpgt_transf_date = dt.datetime.strptime(data_last_2_cpgt, '%d.%m.%Y')
        data_last_3_cpgt = dt.datetime.now().strftime('%d.%m.%Y')
        data_last_3_cpgt_transf_date = dt.datetime.strptime(data_last_3_cpgt, '%d.%m.%Y')
        data_last_1_cpgt_day = data_last_1_1_cpgt_transf_date.day
        data_last_2_cpgt_day = data_last_2_cpgt_transf_date.day
        data_last_3_cpgt_day = data_last_3_cpgt_transf_date.day
        if data_last_3_cpgt_day in range(0, (data_last_2_cpgt_day - data_last_1_cpgt_day)):
            return data_last_1_cpgt.strftime('%d.%m.%Y')
        else:
            return data_last_cpgt.strftime('%d.%m.%Y')

    def data_email(self):
        data_mes_email = dt.datetime.now() + relativedelta(months=1)
        data_email_1 = dt.datetime.now().strftime('01.%m.%Y')
        data_email_1_transf_date = dt.datetime.strptime(data_email_1, '%d.%m.%Y')
        data_email_2 = dt.datetime.now().strftime('21.%m.%Y')
        data_email_2_transf_date = dt.datetime.strptime(data_email_2, '%d.%m.%Y')
        data_email_3 = dt.datetime.now().strftime('%d.%m.%Y')
        data_email_3_transf_date = dt.datetime.strptime(data_email_3, '%d.%m.%Y')
        data_email_1_day = data_email_1_transf_date.day
        data_email_2_day = data_email_2_transf_date.day
        data_email_3_day = data_email_3_transf_date.day
        if data_email_3_day in range(0, (data_email_2_day - data_email_1_day)):
            return dt.datetime.now().strftime('%B')
        else:
            return data_mes_email.strftime('%B')

    def enviar_email(self):
        pergunta_envio = input(
            'Você deseja enviar o email agora?  \n(Pressione "y" para enviar o email ou "enter" para prosseguir no sistema.)')
        if pergunta_envio == 'y':
            data_visivel = self.data_email()

            mes = {'January': 'janeiro',
                   'February': 'fevereiro',
                   'March': 'março',
                   'April': 'abril',
                   'May': 'maio',
                   'Juno': 'junho',
                   'July': 'julho',
                   'August': 'agosto',
                   'Septemper': 'setembro',
                   'October': 'outubro',
                   'November': 'novembro',
                   'December': 'dezembro', }

            data_trad = mes[data_visivel].title()
            smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
            smtpobj.starttls()
            fro = 'jrf.petro@gmail.com'
            to = 'junio_firmino@petrobras.com.br, jrf.petro@gmail.com'

            smtpobj.login(fro, 'yevq kufu ejsx awpz')
            msg = EmailMessage()
            msg['From'] = fro
            msg['To'] = to
            msg['Subject'] = f'Taxas Risco Sacado {data_trad}.'

            msg.set_content(
                f'Prezado Fajardo\n\nSegue as taxas do risco sacado que vigorará em {data_trad}. Peço proceder '
                'com o cadastro. ')
            caminho = open('Cadastro_CPGT_RS(02_07_20).xlsx', 'rb')
            arq_data = caminho.read()
            arq_name = caminho.name
            msg.add_attachment(arq_data, maintype='application', subtype='octet-stream', filename=arq_name)
            smtpobj.send_message(msg)
            smtpobj.quit()
            print('Email enviado!!!!')


x = Risco()
x.interface()
