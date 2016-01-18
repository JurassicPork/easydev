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

import imaplib
import email
from email.parser import BytesParser
from email.header import decode_header
from email.utils import parsedate_tz

from org.universolibre.EasyDev import XLOEMail
from easydev import comun
from easydev.setting import LOG, NAME_EXT, PY2, TIMEOUT

if PY2:
    str = unicode

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


class ImapMail(object):

    def __init__(self, server):
        self.error = ''
        self.con = self._connection(server)
        self._gmail = 'gmail' in server.Name

    def _connection(self, server):
        try:
            if server.Ssl:
                M = imaplib.IMAP4_SSL(server.Name, server.Port)
            else:
                M = imaplib.IMAP4(server.Name, server.Port)
            M.login(server.User, server.Password)
            M.select()
            return M
        except imaplib.IMAP4.error as e:
            self.error = str(e)
            return None
        except Exception as e:
            self.error = str(e)
            return None

    def __del__(self):
        try:
            if self.con:
                self.con.close()
                self.con.logout()
        except:
            pass

    def get_folders(self, subfolders=True):
        SEP = '"."'
        if self._gmail:
            SEP = '"/"'
        folders = ()
        if subfolders:
            typ, subdir = self.con.list()
            sb = [s.decode('utf8').rpartition(SEP)[-1].strip() for s in subdir]
            folders += tuple(sb)
        return folders

    def get_uids(self, folder='INBOX', where='ALL'):
        self.con.select(folder)
        result, data = self.con.uid('search', None, where)
        if result != 'OK':
            return None
        return data[0].decode().split()

    def get(self, uid, only_size=False):
        files = {}
        w = '(RFC822)'
        if only_size:
            w = '(RFC822.SIZE)'
        result, data = self.con.uid('FETCH', uid, w)
        if result != 'OK':
            return {}
        if only_size:
            tmp = data[0].decode().split()
            return {'uid': tmp[2], 'size': int(tmp[4].replace(')', ''))}

        raw = email.message_from_bytes(data[0][1])
        for part in raw.walk():
            #~ print (1, 'Content Type', part.get_content_type())
            cd = part.get('Content-Disposition', None)
            if cd is None or cd == 'inline':
                continue
            if not part.get_content_type() in self.TYPES:
                continue
            file_name = part.get_filename()
            name = file_name[:-4].replace(' ', '_').upper()
            ext = file_name[-3:].lower()
            file_name = '{}.{}'.format(name, ext)
            content = part.get_payload(decode=True)
            files[file_name] = content
        return files

    def _decode_headers(self, headers):
        keys = ('From', 'To', 'Date', 'Subject', 'Content-type')
        data = {'Files': False}
        for k in keys:
            tmp = None
            if headers[k]:
                tmp = decode_header(headers[k])[-1]
                if tmp[1]:
                    tmp = tmp[0].decode(tmp[1])
                else:
                    tmp = tmp[0]
                if isinstance(tmp, bytes):
                    tmp = tmp.decode().strip()
                if k == 'Content-type' and 'multipart/mixed' in tmp:
                    data['Files'] = True
                elif k == 'Date':
                    tmp = parsedate_tz(tmp)
            data[k] = tmp
        return data

    def get_headers(self, uids):
        result, data = self.con.uid(
            'FETCH',
            ','.join(uids),
            '(BODY[HEADER.FIELDS (FROM TO DATE SUBJECT CONTENT-TYPE)])')
        if result != 'OK':
            return None
        data = [
            (r[0].decode().split()[2], BytesParser().parsebytes(r[1], True))
                for r in data if isinstance(r, tuple)
            ]
        return {r[0]: self._decode_headers(r[1]) for r in data}

    def get_all_files(self):
        if self.con is None:
            return self.error
        info = []
        folders = self.get_folders()
        for folder in folders:
            self.con.select(folder)
            info_msg = self.con.status(folder, '(MESSAGES UNSEEN)')
            info.append(str(info_msg))
        return '\n'.join(info)


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

    def readMail(self, server):
        server = ImapMail(server)
        info = server.get_all_files()
        del server
        return info

