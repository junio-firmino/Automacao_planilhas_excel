import datetime as dt
from dateutil.relativedelta import relativedelta


def data_cadastro():
    data_cad = dt.datetime.now()
    return data_cad.strftime('%d.%m.%Y')

def data_cadastro_month():
    data_cad = dt.datetime.now()
    return data_cad.strftime('.%m.%Y')

def data_inicio():
    data_ini = dt.datetime.now() + relativedelta(months=1)
    data_1 = dt.datetime.now().strftime('01.%m.%Y')
    data_1_transforma_date = dt.datetime.strptime(data_1, '%d.%m.%Y')
    data_2 = dt.datetime.now().strftime('21.%m.%Y')
    data_2_transforma_date = dt.datetime.strptime(data_2, '%d.%m.%Y')
    data_3 = dt.datetime.now().strftime('%d.%m.%Y')
    data_3_transforma_date = dt.datetime.strptime(data_3, '%d.%m.%Y')
    data_1_day = data_1_transforma_date.day
    data_2_day = data_2_transforma_date.day
    data_3_day = data_3_transforma_date.day
    if data_3_day in range(0, (data_2_day - data_1_day)):
        return data_cadastro()
    else:
        return data_ini.strftime('01.%m.%Y')

def data_last_day_risco_sacado():
    data_last = dt.datetime.now() + relativedelta(day=31, months=1)
    data_last_1 = dt.datetime.now() + relativedelta(day=31)
    data_last_1_1 = dt.datetime.now().strftime('01.%m.%Y')
    data_last_1_1_transforma_date = dt.datetime.strptime(data_last_1_1, '%d.%m.%Y')
    data_last_2 = dt.datetime.now().strftime('21.%m.%Y')
    data_last_2_transforma_date = dt.datetime.strptime(data_last_2, '%d.%m.%Y')
    data_last_3 = dt.datetime.now().strftime('%d.%m.%Y')
    data_last_3_transforma_date = dt.datetime.strptime(data_last_3, '%d.%m.%Y')
    data_last_1_day = data_last_1_1_transforma_date.day
    data_last_2_day = data_last_2_transforma_date.day
    data_last_3_day = data_last_3_transforma_date.day
    if data_last_3_day in range(0, (data_last_2_day - data_last_1_day)):
        return data_last_1.strftime('%d.%m.%Y')
    else:
        return data_last.strftime('%d.%m.%Y')


def data_save_arquivo():
    data_save = dt.datetime.now()
    return data_save.strftime('%d_%m_%y')


def data_last_day_cpgt():
    data_last_cpgt = dt.datetime.now() + relativedelta(day=1, months=3)
    data_last_1_cpgt = dt.datetime.now() + relativedelta(day=1, months=2)
    data_last_1_1_cpgt = dt.datetime.now().strftime('01.%m.%Y')
    data_last_1_1_cpgt_transforma_date = dt.datetime.strptime(data_last_1_1_cpgt, '%d.%m.%Y')
    data_last_2_cpgt = dt.datetime.now().strftime('21.%m.%Y')
    data_last_2_cpgt_transforma_date = dt.datetime.strptime(data_last_2_cpgt, '%d.%m.%Y')
    data_last_3_cpgt = dt.datetime.now().strftime('%d.%m.%Y')
    data_last_3_cpgt_transforma_date = dt.datetime.strptime(data_last_3_cpgt, '%d.%m.%Y')
    data_last_1_cpgt_day = data_last_1_1_cpgt_transforma_date.day
    data_last_2_cpgt_day = data_last_2_cpgt_transforma_date.day
    data_last_3_cpgt_day = data_last_3_cpgt_transforma_date.day
    if data_last_3_cpgt_day in range(0, (data_last_2_cpgt_day - data_last_1_cpgt_day)):
        return data_last_1_cpgt.strftime('%d.%m.%Y')
    else:
        return data_last_cpgt.strftime('%d.%m.%Y')


def data_email():
    data_mes_email = dt.datetime.now() + relativedelta(months=1)
    data_email_1 = dt.datetime.now().strftime('01.%m.%Y')
    data_email_1_transforma_date = dt.datetime.strptime(data_email_1, '%d.%m.%Y')
    data_email_2 = dt.datetime.now().strftime('21.%m.%Y')
    data_email_2_transforma_date = dt.datetime.strptime(data_email_2, '%d.%m.%Y')
    data_email_3 = dt.datetime.now().strftime('%d.%m.%Y')
    data_email_3_transforma_date = dt.datetime.strptime(data_email_3, '%d.%m.%Y')
    data_email_1_day = data_email_1_transforma_date.day
    data_email_2_day = data_email_2_transforma_date.day
    data_email_3_day = data_email_3_transforma_date.day
    if data_email_3_day in range(0, (data_email_2_day - data_email_1_day)):
        return dt.datetime.now().strftime('%B')
    else:
        return data_mes_email.strftime('%B')

def data_inicio_ant():
    data_ini_ant = dt.datetime.now() + relativedelta(months=-1)
    data_1_ini_ant = dt.datetime.now().strftime('01.%m.%Y')
    data_1_transforma_date_ini_ant = dt.datetime.strptime(data_1_ini_ant, '%d.%m.%Y')
    data_2_ini_ant = dt.datetime.now().strftime('21.%m.%Y')
    data_2_transforma_date_ini_ant = dt.datetime.strptime(data_2_ini_ant, '%d.%m.%Y')
    data_3_ini_ant = dt.datetime.now().strftime('%d.%m.%Y')
    data_3_transforma_date_ini_ant = dt.datetime.strptime(data_3_ini_ant, '%d.%m.%Y')
    data_1_day_ini_ant = data_1_transforma_date_ini_ant.day
    data_2_day_ini_ant = data_2_transforma_date_ini_ant.day
    data_3_day_ini_ant = data_3_transforma_date_ini_ant.day
    # if data_3_day_ini_ant in range(0, (data_2_day_ini_ant - data_1_day_ini_ant)):
    #     return data_cadastro()
    # else:
    return data_ini_ant.strftime('01.%m.%Y')



x= data_inicio_ant()
print(x)