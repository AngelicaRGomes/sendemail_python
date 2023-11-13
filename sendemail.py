from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from datetime import datetime
import pytesseract
import smtplib
import cv2
import re

# Email sending time
now = datetime.now()  # current date and time
time = now.strftime("%H")
print("time:", time)

email = 'robot-test@outlook.com'
password = ''
send_to_email = ['angelicarodrigues@examplehost.com', 'angellicarodrigues@examplehost.com']


def sendemail():
    # Email configuration
    subject = f'Carregamento diário das {time}h - Telesena INCs'
    message = f'''
        <html>
        <head></head>
        <body>
            <p style="font-size:18px;">Bom Dia! Time, segue o carregamento diário das {time} horas para incidentes.</p>
            <img src="cid:C:\\Users\\011391631\\Downloads\\INCs_fluxo_SLO.png" alt="planilha"/>
            <p style="font-size:18px;">Backlog SLO INC</p>
            <img src="cid:C:\\Users\\011391631\\Downloads\\INCs_backlog_SLO.png" alt="planilha"/>
        </body>
        </html>
        '''

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = ", ".join(send_to_email)
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'html'))

    # Setup the attachment

    files = ['C:\\Users\\011391631\\Downloads\\INCs_fluxo_SLO.png', 'C:\\Users\\011391631\\Downloads\\INCs_backlog_SLO.png']

    for file in files:
        fp = open(file, 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()
        file_name = fp.name
        msgImage.add_header('Content-ID', file_name)
        msg.attach(msgImage)

    try:
        server = smtplib.SMTP('SMTP.office365.com', 587)
        server.starttls()
        server.login(email, password)
        text = msg.as_string()
        server.sendmail(email, send_to_email, text)

        print('\nE-mail sent correctly!')

    except Exception as e:
        print("\nError trying to send the e-mail", e)
    finally:
        print("\nClosing the server...")
        server.quit()


sendemail()
