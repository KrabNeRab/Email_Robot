import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart

import openpyxl
import pyodbc
import smtplib

import schedule

import sql_queries
import config

class DataBaseConnector:
    def __init__(self, server, database):
        self.server = server
        self.database = database

    def connectioin(self):
        conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+self.server+';DATABASE='+self.database+';Trusted_Connection=yes')
        return conn

class ExcelGeneration:
    def __init__(self, data, cursor):
        self.data = data
        self.cursor = cursor

    def generation_excel(self):

        wb = openpyxl.Workbook()
        sheet = wb.active

        #добавляем заголовки
        headers = [desc[0] for desc in self.cursor.description]
        sheet.append(headers)

        #Добавляем данные с базы
        for row in self.data:
            sheet.append(list(row))
        wb.save('выгрузка_по_отправлениям_l-post.xlsx')
        return 'выгрузка_по_отправлениям_l-post.xlsx'



class EmailSender:
    def __init__(self, email_sender, email_recipient, subject,  attachment, smtp_server, smtp_port, email_password):
        self.email_sender = email_sender
        self.email_recipient = email_recipient
        self.subject = subject
        self.attachment = attachment
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_password = email_password


    def send_email(self):

        msg = MIMEMultipart()

        msg['From'] = self.email_sender
        msg['To'] = self.email_recipient
        msg['Subject'] = self.subject


        with open(self.attachment, 'rb') as f:
            self.attachment = MIMEApplication(f.read(), Name=self.attachment)
        msg.attach(self.attachment)
        server = smtplib.SMTP(self.smtp_server, self.smtp_port)
        server.login(self.email_sender, self.email_password)
        server.sendmail(self.email_sender, self.email_recipient, msg.as_string())
        server.quit()


def send_email():
    connector = DataBaseConnector(config.sql_server, config.DataBase)
    conn = connector.connectioin()
    cursor = conn.cursor()
    cursor.execute(sql_queries.sql_query)
    data = cursor.fetchall()


    generator = ExcelGeneration(data, cursor)
    attachment = generator.generation_excel()

    sender = EmailSender(config.email_sender, sql_queries.email_recipient, sql_queries.email_subject, attachment, config.smtp_server, config.smtp_port, config.email_password)
    sender.send_email()



#раз в минуту будет отправляться письмо
def testing():
    schedule.every(1).minutes.do(send_email)

testing()

while True:
    schedule.run_pending()
    time.sleep(1)






