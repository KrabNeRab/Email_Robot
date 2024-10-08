всю инфу тянем из файлов config и sql_queries
config: smtp_server, smtp_port, sql_server, DataBase, email_sender, email_password
sql_queries: sql_query, email_recipient, email_subject



если несколько адрессатов, то msg['To'] = ', '.join(self.email_recipient)
в sql_queries email_recipient = ['email1@mail.ru', 'email2@mail.ru']
