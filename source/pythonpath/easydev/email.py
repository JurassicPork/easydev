# coding: utf-8

import logging
from threading import Thread
import smtplib
from smtplib import SMTPException, SMTPAuthenticationError
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from org.universolibre.EasyDev import XEMail
from easydev.setting import LOG, NAME_EXT, TIMEOUT


log = logging.getLogger(NAME_EXT)


def send_mail(server, message):
    sender = server.User
    receivers = message.To.split(',') + message.Cc.split(',') + message.Bcc.split(',')

    email = MIMEMultipart()
    email['From'] = sender
    email['To'] = message.To
    email['Cc'] = message.Cc
    email['Subject'] = message.Subject
    body = message.Body.replace('\\n', '<br />').replace('\n', '<br />')
    email.attach(MIMEText(body, 'html'))

    for f in message.Files:
        path = path_os(f)
        if not os.path.exists(path):
            continue
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(path, 'rb').read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            "attachment; filename={}".format(os.path.basename(path)))
        email.attach(part)
    try:
        smtp = smtplib.SMTP(server.Name, server.Port, timeout=TIMEOUT)
        if server.Ssl:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
        log.info('Connect server...')
        smtp.login(server.User, server.Password)
        log.info('Send mail...')
        smtp.sendmail(sender, receivers, email.as_string())
        log.info('Log out server...')
        smtp.quit()
        log.info('Close...')
        return True
    except SMTPAuthenticationError as e:
        if e[0] == 534 and 'gmail' in server.Server:
            msg = 'Necesitas activar el acceso a otras aplicaciones en tu cuenta de GMail'
            log.debug(msg)
            return False
        elif e[0] == 535:
            msg = 'Nombre de usuario o contraseña inválidos'
            log.debug(msg)
            return False
    except SMTPException as e:
        log.debug(e)
        return False
    except Exception as e:
        log.debug(e)
        return False


class EMail(XEMail):

    def __init__(self):
        pass

    def sendMail(self, server, message):
        if not server.Port:
            server.Port = 587
        if server.Thread:
            log.info('Send in thread...')
            thread = Thread(target=send_mail, args=(server, message))
            thread.start()
            thread.join()
        else:
            log.info('Send normal...')
            return send_mail(server, message)

