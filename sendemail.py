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


def sendemail_incs():
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


def sendemail_ritms():
    # Email configuration
    subject = f'Carregamento diário das {time}h - Telesena RITMs'
    message = f'''
        <html>
        <head></head>
        <body>
            <p style="font-size:18px;">Bom Dia! Time, segue o carregamento diário das {time} horas para requisições.</p>
            <img src="cid:C:\\Users\\011391631\\Downloads\\RITMs_fluxo_SLO.png" alt="planilha"/>
            <p style="font-size:18px;">Backlog SLO RITMs</p>
            <img src="cid:C:\\Users\\011391631\\Downloads\\RITMS_backlog_SLO.png" alt="planilha"/>
        </body>
        </html>
        '''

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = ", ".join(send_to_email)
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'html'))

    # Setup the attachment

    files = ['C:\\Users\\011391631\\Downloads\\RITMs_fluxo_SLO.png', 'C:\\Users\\011391631\\Downloads\\RITMS_backlog_SLO.png']

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


def sendemail_sem_torre_incs():
    # Email configuration
    send_to_email_dpe = ['angelicarodrigues@examplehost.com', '<angellyka_lly@examplehost.com>']
    subject = f'Sem Torre maior que 0 das {time}h - Telesena INCs'
    message = f'''
        <html>
        <head></head>
        <body>
            <p style="font-size:18px;">Bom Dia! Sem Torre consta maior que 0 em INCs, novo serviço em espera.</p>
            <img src="cid:C:\\Users\\011391631\\Downloads\\sem_torre_incs.png" alt="planilha"/>
        </body>
        </html>
        '''

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = ", ".join(send_to_email_dpe)
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'html'))

    # Setup the attachment

    file = 'C:\\Users\\011391631\\Downloads\\sem_torre_incs.png'
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
        server.sendmail(email, send_to_email_dpe, text)

        print('\nE-mail sent correctly!')

    except Exception as e:
        print("\nError trying to send the e-mail", e)
    finally:
        print("\nClosing the server...")
        server.quit()


def sendemail_sem_torre_ritms():
    # Email configuration
    send_to_email_dpe = ['angelicarodrigues@examplehost.com', 'angellica@examplehost.com']
    subject = f'Sem Torre maior que 0 das {time}h - Telesena RITMs'
    message = f'''
        <html>
        <head></head>
        <body>
            <p style="font-size:18px;">Bom Dia! Sem Torre consta maior que 0 em RITMs, novo serviço em espera.</p>
            <img src="cid:C:\\Users\\011391631\\Downloads\\sem_torre_ritms.png" alt="planilha"/>
        </body>
        </html>
        '''

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = ", ".join(send_to_email_dpe)
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'html'))

    # Setup the attachment

    file = 'C:\\Users\\011391631\\Downloads\\sem_torre_ritms.png'
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
        server.sendmail(email, send_to_email_dpe, text)

        print('\nE-mail sent correctly!')

    except Exception as e:
        print("\nError trying to send the e-mail", e)
    finally:
        print("\nClosing the server...")
        server.quit()


def img_to_string_incs():
    name_img = "C:\\Users\\011391631\\Downloads\\sem_torre_incs.png"
    img = cv2.imread(name_img)
    result = pytesseract.image_to_string(img)
    result_space = re.sub("\s", "", result)
    if re.search('semtorre:0', result_space, re.IGNORECASE):
        print("Found")
        print(result_space)
    else:
        print("Not Found")
        sendemail_sem_torre_incs()



def img_to_string_ritms():
    name_img2 = "C:\\Users\\011391631\\Downloads\\sem_torre_ritms.png"
    img2 = cv2.imread(name_img2)
    result2 = pytesseract.image_to_string(img2)
    result2_space = re.sub("\s", "", result2)
    if re.search('semtorre:0', result2_space, re.IGNORECASE):
        print("Found")
        print(result2_space)
    else:
        print("Not Found")
        sendemail_sem_torre_ritms()
