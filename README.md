всю инфу тянем из файлов config и sql_queries

sql_queries:
1)sql_query (запрос в бд)
2)email_recipient (адресат)
3)email_subject (Тема письма)
4)file_name (название файла)

сonfig:
1)smtp_server
2)smtp_port
3)sql_server
4)DataBase (название бд)
5)email_sender (отправитель)
6)email_password (пароль от почты)


если несколько адрессатов, то msg['To'] = ', '.join(self.email_recipient)
в sql_queries email_recipient = ['email1@mail.ru', 'email2@mail.ru']
