
'''This code sends personalized automated emails with attachments,
using Gmail, getting data from an excel file.'''


import smtplib
import email.message
import pandas as pd
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

excel_data = pd.read_excel('send_e-mails.xlsx', sheet_name='1')

def send_email():
    for i, recipient_email in enumerate(excel_data['E-mail']):
        manager = excel_data.loc[i, 'Manager']
        area = excel_data.loc[i, 'Report']
        email_body = """
        <p>Hi, {}!</p>
        <p>Here goes your report. Should you have any questions, do not hesite to contact me .</p>
        <p>Best regards,</p>
        <p>Pedro Oliveira</p>
        </p>Head of technology</p>
        """.format(manager, area)

        mail = MIMEMultipart()
        mail['Subject'] = "Automated Report - {}.".format(area)
        mail['From'] = 'peterhenrike@gmail.com'
        mail['To'] = recipient_email

        password = 'pfsunxuwgmzhilbb' 

        mail.attach(MIMEText(email_body, 'html'))

        mime_type, _ = mimetypes.guess_type('/Users/pedro/Desktop/Python_Impressionador/Desafio_email /{}.xlsx'.format(area))
        with open('/Users/pedro/Desktop/Python_Impressionador/Desafio_email /{}.xlsx'.format(area), 'rb') as fp:
            dados = fp.read()
        attachment = MIMEBase(mime_type.split("/")[0],mime_type.split("/")[1])
        attachment.set_payload(dados)
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment',filename="{}.xlsx".format(area))
        mail.attach(attachment)

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        s.login(mail['From'], password)
        s.sendmail(mail['From'], [mail['To']], mail.as_string().encode('utf-8'))
        print('Email sent')

send_email()