import smtplib
import pandas as pd
import pathlib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#function to send email by gmail
def send_email():  

    msg = MIMEMultipart()
    msg['Subject'] = 'Promotions of the day'
    msg['From'] = 'agathabarros@gmail.com'
    msg['To'] = 'exemplo@gmail.com'
    password = 'sua senha' 


        #mail body HTML
    mail_body = f'''
        <p>Good Morning</p>
    '''

    msg.attach(MIMEText(mail_body, 'html'))


    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email sent')



send_email()