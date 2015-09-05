# coding: utf-8

import logging
from threading import Thread
import smtplib
from smtplib import SMTPException, SMTPAuthenticationError
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from org.universolibre.util.EasyDev import XEMail
from easydev.setting import LOG, NAME_EXT


log = logging.getLogger(NAME_EXT)


def send_mail(server, mail, files):
    mail = {r[0]: r[1] for r in mail}
    sender = server['smtpUser']
    receivers = mail['to'].split(',')

    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = mail['to']
    message['Cc'] = mail.get('cc', '')
    message['Subject'] = mail['subject']
    body = mail['body'].replace('\\n', '<br />').replace('\n', '<br />')
    message.attach(MIMEText(body, 'html'))
    for f in files:
        path = path_os(f)
        if not os.path.exists(path):
            continue
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(path, 'rb').read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            "attachment; filename={}".format(os.path.basename(path)))
        message.attach(part)
    if message['Cc']:
        receivers += mail['cc'].split(',')
    if mail.get('bcc', ''):
        receivers += mail['bcc'].split(',')
    try:
        smtp = smtplib.SMTP(server['smtpServer'], server['smtpPort'], timeout=10)
        if server['smtpSsl']:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
        log.info('Connect server...')
        smtp.login(server['smtpUser'], server['smtpPassword'])
        log.info('Send mail...')
        smtp.sendmail(sender, receivers, message.as_string())
        log.info('Log out server...')
        smtp.quit()
        log.info('Close...')
        return True
    except SMTPAuthenticationError as e:
        if e[0] == 534 and 'gmail' in server['smtpServer']:
            msg = 'Necesitas activar el acceso a otras aplicaciones en tu cuenta de GMail'
            log.debug(msg)
            return False
        elif e[0] == 535:
            msg = 'Nombre de usuario o contraseña inválidos'
            log.debug(msg)
            return False
    except SMTPException as e:
        log.debug(e, exc_info=True)
        return False
    except Exception as e:
        log.debug(e, exc_info=True)
        return False


class EMail(XEMail):
    smtpServer = ''
    smtpUser = ''
    smtpPassword = ''

    def __init__(self):
        self.smtpPort = 587
        self.smtpSsl = True
        self.smtpThread = True

    def sendMail(self, mail, files):
        if self.smtpThread:
            log.info('Send in thread...')
            thread = Thread(target=send_mail, args=(self.__dict__, mail, files))
            thread.start()
            thread.join()
        else:
            log.info('Send normal...')
            return send_mail(self.__dict__, mail, files)

