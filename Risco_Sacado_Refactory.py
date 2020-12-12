from openpyxl import load_workbook
import smtplib
from email.message import EmailMessage
from locale import setlocale, LC_ALL
import assists

# I will refactory This project, for do it I chose the design pattern Facade.