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

    def connection(self):
        conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+self.server+';DATABASE='+self.database+';Trusted_Connection=yes')
        return conn

class ExcelGeneration:
    def __init__(self, data, cursor, file_name):
        self.data = data
        self.cursor = cursor
        self.file_name = file_name
    
    def generation_excel(self):

        wb = openpyxl.Workbook()
        sheet = wb.active

        #добавляем заголовки
        headers = [desc[0] for desc in self.cursor.description]
        sheet.append(headers)

        #Добавляем данные с базы
        for row in self.data:
            sheet.append(list(row))
        wb.save(self.file_name)
        return self.file_name



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


class NewsLetters:
    def __init__(self, sql_server, DataBase, sql_query, email_sender, email_recipient, email_subject, smtp_server, smtp_port, email_password, file_name):
        self.sql_server = sql_server
        self.DataBase = DataBase
        self.sql_query = sql_query
        self.email_sender = email_sender
        self.email_recipient = email_recipient
        self.email_subject = email_subject
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_password = email_password
        self.file_name = file_name

    def start(self):
        connector = DataBaseConnector(self.sql_server, self.DataBase)
        conn = connector.connection()
        cursor = conn.cursor()
        cursor.execute(self.sql_query)
        data = cursor.fetchall()

        generator = ExcelGeneration(data, cursor, self.file_name)
        attachment = generator.generation_excel()

        sender = EmailSender(self.email_sender, self.email_recipient, self.email_subject, attachment, self.smtp_server, self.smtp_port, self.email_password)
        sender.send_email()


email = NewsLetters(config.sql_server, config.DataBase, sql_queries.sql_query, config.email_sender, sql_queries.email_recipient, sql_queries.email_subject, config.smtp_server, config.smtp_port, config.email_password, sql_queries.file_name)
email.start()


while True:
    schedule.run_pending()
    time.sleep(1)





