import os

import smtplib
from email.message import EmailMessage

#EMAIL_ADDRESS = os.environ.get('EMAIL_USER')
#EMAIL_PASSWORD = os.environ.get('EMAIL_PASS')



class Mailing:

    def __init__(self):
        """Define constant email contents"""
        self.msg = EmailMessage()

        self.msg['From'] = os.environ.get('EMAIL_USER')
        self.msg['To'] = ['dariusz.giemza@axi-card.pl','mateusz.filipiak@axi-card.pl','anna.rzemek@axi-card.pl']

    def send_critical_message(self,content):

        self.msg['Subject'] = "FINES Reporting Error"
        self.msg.set_content(content)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(os.environ.get('EMAIL_USER'), os.environ.get('EMAIL_PASSWORD'))

            smtp.send_message(self.msg)
            smtp.close()

    def send_success_message(self):

        self.msg['Subject'] = "FINES Successfully reporting"
        self.msg.set_content("Report has been successfully prepared")

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(os.environ.get('EMAIL_USER'), os.environ.get('EMAIL_PASSWORD'))

            smtp.send_message(self.msg)
            smtp.close()