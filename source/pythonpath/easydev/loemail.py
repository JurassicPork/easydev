# coding: utf-8

import logging
from threading import Thread
import smtplib
from smtplib import SMTPException, SMTPAuthenticationError
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import mailbox

from org.universolibre.EasyDev import XLOEMail
from easydev import comun
from easydev.setting import LOG, NAME_EXT, TIMEOUT


log = logging.getLogger(NAME_EXT)


def save_message(path, email):
    mbox = mailbox.mbox(path, create=True)
    mbox.lock()
    try:
        msg = mailbox.mboxMessage(email)
        mbox.add(msg)
        mbox.flush()
    finally:
        mbox.unlock()
    return


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
        path = comun.path_to_os(f)
        if not comun.exists(path):
            continue
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(path, 'rb').read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            "attachment; filename={}".format(comun.basename(path)))
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
        if server.PathSave and message.Save:
            save_message(server.PathSave, email)
        log.info('Close...')
        return True
    except SMTPAuthenticationError as e:
        err = str(e)
        if '534' in err and 'gmail' in server.Server:
            msg = 'Need activate cant access other app in your account GMail'
            log.debug(msg)
            return False
        elif '535' in err:
            msg = 'User or password are invalid'
            log.debug(msg)
            return False
    except SMTPException as e:
        log.debug(e)
        return False
    except Exception as e:
        log.debug(e)
        return False


class LOEMail(XLOEMail):

    def __init__(self):
        pass

    def sendMail(self, server, message):
        if not server.Port:
            server.Port = 587
        if server.Thread:
            log.info('Send in thread...')
            thread = Thread(target=send_mail, args=(server, message))
            thread.start()
            return True
        else:
            log.info('Send normal...')
            return send_mail(server, message)

    def readMail(self, server, all_msg=False):
        return

