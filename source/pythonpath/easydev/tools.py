# coding: utf-8

import ctypes
import subprocess
import sys
import os
import re
import getpass
import platform
import json
import logging
import smtplib
from smtplib import SMTPException, SMTPAuthenticationError
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pprint import pprint
from string import Template

import uno
from com.sun.star.beans import PropertyValue

try:
    from setting import DESKTOP, OS, WIN, WRITER, TOOLKIT, EXT_PDF, NODE, \
        NODE_CONFIG, LOG, NAME_EXT, BUTTONS_OK, BUTTONS_YES_NO, YES
except:
    from easydev.setting import DESKTOP, OS, WIN, WRITER, TOOLKIT, EXT_PDF, \
        NODE, NODE_CONFIG, LOG, NAME_EXT, BUTTONS_OK, BUTTONS_YES_NO, YES


log = logging.getLogger(NAME_EXT)
CTX = uno.getComponentContext()
SM = CTX.getServiceManager()


class OutputDoc(object):

    def __init__(self, doc):
        self.doc = doc

    def write(self, info):
        text = self.doc.Text
        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        text.insertString(cursor, str(info), 0)
        return


def _create_instance(name=DESKTOP, with_context=True):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance


def _make_properties(properties):
    prop = []
    l = len(properties)
    for i in range(0, l, 2):
        pv = PropertyValue()
        pv.Name = properties[i]
        pv.Value = properties[i + 1]
        prop.append(pv)
    return tuple(prop)


def debug(data):
    """ Show data for debug
        If SO is Win, show data in Writer document
        else, show data in stdout
    """
    if OS == WIN:
        doc = get_doc('debug.odt')
        if not doc:
            doc = new_doc(WRITER)
        out = OutputDoc(doc)
        sys.stdout = out
    pprint (data)
    return


def msgbox(message, type_msg='infobox', title='Debug', buttons=BUTTONS_OK):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
    """
    desktop = _create_instance()
    toolkit = _create_instance(TOOLKIT, False)
    parent = toolkit.getDesktopWindow()
    mb = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return mb.execute()


def question(title, message):
    return YES == msgbox(message, 'querybox', title, BUTTONS_YES_NO)


def cmd(command, *data):
    """
        Execute methods by name
    """
    return globals()[command](*data)


def new_doc(type_doc):
    """
        Create new doc
        http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XComponentLoader.html

    type_doc:
        scalc
        swriter
        simpress
        sdraw
        smath
    """
    if not type_doc:
        type_doc = 'scalc'
    desktop = _create_instance()
    path = 'private:factory/{}'.format(type_doc)
    doc = desktop.loadComponentFromURL(path, '_default', 0, ())
    return doc


def get_doc(title=''):
    """
        If title is missing get current component,
        else search doc title in components
    """
    desktop = _create_instance()
    if not title:
        return desktop.getCurrentComponent()

    enum = desktop.getComponents().createEnumeration()
    while enum.hasMoreElements():
        doc = enum.nextElement()
        if doc.getTitle() == title:
            return doc
    return None


def get_type_doc(doc):
    """
        Get type doc
    """
    services = {
        'calc': 'com.sun.star.sheet.SpreadsheetDocument',
        'writer': 'com.sun.star.text.TextDocument',
        'impress': 'com.sun.star.presentation.PresentationDocument',
        'draw': 'com.sun.star.drawing.DrawingDocument',
        'math': 'com.sun.star.formula.FormulaProperties',
        'base': 'com.sun.star.sdb.OfficeDatabaseDocument',
        'ide': 'com.sun.star.script.BasicIDE',
    }
    for key, value in services.items():
        if doc.supportsService(value):
            return key
    return ""


def get_docs():
    """
        Return all documents open
    """
    docs = []
    desktop = _create_instance()
    enum = desktop.getComponents().createEnumeration()
    while enum.hasMoreElements():
        docs.append(enum.nextElement())
    return tuple(docs)


def open_doc(path, options):
    """
        Open doc
        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html
        http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1MediaDescriptor.html
    """
    properties = _make_properties(options)
    path_url = path_to_url(path)
    desktop = _create_instance()
    doc = desktop.loadComponentFromURL(path_url, '_blank', 0, properties)
    return doc


def set_focus(doc):
    """
        Active doc
    """
    window = doc.getCurrentController().getFrame().getComponentWindow()
    window.setFocus()
    return


def get_status_bar(doc):
    """
        Return status bar
    """
    statusbar = doc.getCurrentController().getStatusIndicator()
    return statusbar


def export_pdf(doc, path_save, options):
    """
        Export to PDF
        http://wiki.services.openoffice.org/wiki/API/Tutorials/PDF_export
    """
    close = False
    if isinstance(doc, str):
        close = True
        doc = open_doc(doc, ('Hidden', True))
    if path_save:
        if os.path.isdir(path_save):
            _, _, name, extension = get_path_info(path_os(doc.getURL()))
            path_save = path_to_url(
                os.path.normpath('{}/{}.{}'.format(path_save, name, EXT_PDF)))
        else:
            path_save = path_to_url(path_save)
    else:
        path_save = path_to_url(replace_ext(path_os(doc.getURL()), EXT_PDF))
    type_doc = get_type_doc(doc)
    filters = {
        'calc': 'calc_pdf_Export',
        'writer': 'writer_pdf_Export',
        'impress': 'impress_pdf_Export',
        'draw': 'draw_pdf_Export',
        'math': 'math_pdf_Export',
    }
    filter_data = _make_properties(options)
    media_descriptor = _make_properties((
        'FilterName', filters[type_doc],
        'FilterData', uno.Any("[]com.sun.star.beans.PropertyValue", filter_data)
    ))
    doc.storeToURL(path_save, media_descriptor)
    if close:
        doc.dispose()
    if os.path.exists(path_os(path_save)):
        return path_save
    return ''


def array(array, method, data):
    """
        Methods of list to Basic
    """
    res = None
    l = list(array)
    if method == 'insert':
        res = getattr(l, method)(*data)
    elif method == 'pop':
        res = getattr(l, method)(data)
        res = (tuple(l), res)
    elif method == 'remove_all':
        l = [i for i in array if i != data]
    elif method in ('reverse', 'sort'):
        res = getattr(l, method)()
    elif method == 'unique':
        l = list(set(l))
    elif method in ('len', 'max', 'min'):
        res = eval('{}({})'.format(method, l))
    elif method == 'slice':
        l = eval('{}{}'.format(l, data))
    elif method == 'in':
        res = data in l
    else:
        res = getattr(l, method)(data)
    if res is None:
        return tuple(l)
    else:
        return res


def get_size_screen():
    if OS == WIN:
        user32 = ctypes.windll.user32
        res = '{}x{}'.format(user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
    else:
        args = 'xrandr | grep "\*" | cut -d" " -f4'
        res = subprocess.check_output(args, shell=True).decode()
        return res


def get_info_pc():
    """
        Get info PC:
        name user,
        name pc,
        system/OS name,
        machine type,
        Returns the (real) processor name
        string identifying platform with as much useful information as possible,
    """
    info = (
        getpass.getuser(),
        platform.node(),
        platform.system(),
        platform.machine(),
        platform.processor(),
        platform.platform(),
    )
    return info


def path_to_url(path):
    if path.startswith('file://'):
        return path
    return uno.systemPathToFileUrl(path)


def path_os(path):
    if path.startswith('file://'):
        path = uno.fileUrlToSystemPath(path)
    return path


def get_path(name):
    """
        Return de path name in config
        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XPathSettings.html
    """
    if not name:
        name = 'Work'
    path = _create_instance('com.sun.star.util.PathSettings')
    return getattr(path, name)


def get_path_info(path):
    path, filename = os.path.split(path)
    name, extension = os.path.splitext(filename)
    return (path, filename, name, extension)


def replace_ext(path, ext):
    path, _, name, _ = get_path_info(path)
    return '{}/{}.{}'.format(path, name, ext)


def path_join(paths):
    return os.path.normpath(os.path.join(*paths))


def get_folder(init_folder=''):
    if init_folder:
        init_folder = path_to_url(init_folder)
    else:
        init_folder = get_path('Work')
    folder = _create_instance('com.sun.star.ui.dialogs.FolderPicker')
    folder.setDisplayDirectory(init_folder)
    if folder.execute():
        return folder.getDirectory()
    else:
        return ''


def get_selected_files(init_folder, multiple, filters):
    if init_folder:
        init_folder = path_to_url(init_folder)
    else:
        init_folder = get_path('Work')

    folder = _create_instance('com.sun.star.ui.dialogs.FilePicker')
    folder.setDisplayDirectory(init_folder)
    folder.setMultiSelectionMode(multiple)
    if filters:
        folder.setCurrentFilter(filters[0])
        for i in range(0, len(filters), 2):
            folder.appendFilter(filters[i], filters[i + 1])

    if folder.execute():
        files = folder.getSelectedFiles()
        if multiple:
            return files
        else:
            return files[0]
    else:
        return ""


def get_files(path, ext):
    paths = []
    for folder, _, files in os.walk(path):
        pattern = re.compile('\.{}'.format(ext), re.IGNORECASE)
        paths += [os.path.join(folder, f) for f in files if pattern.search(f)]
    return tuple(paths)


def execute(args, wait):
    if wait:
        res = subprocess.check_output(' '.join(args), shell=True).decode()
        return res
    else:
        subprocess.Popen(args)
    return


def get_config(key):
    name = 'com.sun.star.configuration.ConfigurationProvider'
    cp = _create_instance(name)
    node = PropertyValue()
    node.Name = 'nodepath'
    node.Value = NODE
    try:
        ca = cp.createInstanceWithArguments(
            'com.sun.star.configuration.ConfigurationAccess', (node,))
        if ca and (ca.hasByName(NODE_CONFIG)):
            data = ca.getPropertyValue(NODE_CONFIG)
            if not data:
                return data
            data = json.loads(data)
            if key:
                value = data.get(key, '')
                if isinstance(value, list):
                    return tuple(value)
                else:
                    return value
            else:
                return data
        return
    except Exception as e:
        log.debug(e, exc_info=True)
        return


def set_config(key, value):
    name = 'com.sun.star.configuration.ConfigurationProvider'
    cp = _create_instance(name)
    node = PropertyValue()
    node.Name = 'nodepath'
    node.Value = NODE
    try:
        config_writer = cp.createInstanceWithArguments(
            'com.sun.star.configuration.ConfigurationUpdateAccess', (node,))
        data = get_config('')
        if not data:
            data = {}
        data[key] = value
        new_values = json.dumps(data)
        config_writer.setPropertyValue(NODE_CONFIG, new_values)
        config_writer.commitChanges()
        return True
    except Exception as e:
        log.debug(e, exc_info=True)
        return False


def send_mail(server, mail, files):
    # server, port, ssl, user, pass
    server = {r[0]: r[1] for r in server}
    mail = {r[0]: r[1] for r in mail}
    sender = server['user']
    receivers = mail['to'].split(',')

    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = mail['to']
    message['Cc'] = mail.get('cc', '')
    message['Bcc'] = mail.get('bcc', '')
    message['Subject'] = mail['subject']
    message.attach(MIMEText(mail['body'], 'html'))
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
    if message['Bcc']:
        receivers += mail['bcc'].split(',')
    try:
        smtp = smtplib.SMTP(server['server'], server['port'], timeout=10)
        if server['ssl']:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
        log.info('Connect server...')
        smtp.login(server['user'], server['pass'])
        log.info('Send mail...')
        smtp.sendmail(sender, receivers, message.as_string())
        log.info('Log out server...')
        smtp.quit()
        log.info('Close...')
        return True
    except SMTPAuthenticationError as e:
        if e[0] == 534 and 'gmail' in server['server']:
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


def render(template, data):
    data = {r[0]: r[1] for r in data}
    s = Template(template)
    return s.safe_substitute(**data)


def file_open(path, mode='r'):
    data = ''
    if not mode:
        mode = 'r'
    path = path_os(path)
    with open(path, mode) as f:
        data = f.read()
    return data
