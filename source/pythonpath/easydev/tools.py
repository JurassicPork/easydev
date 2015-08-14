# coding: utf-8
import uno
import ctypes
import subprocess
import sys
import os
import re
import getpass
import platform
from pprint import pprint

from com.sun.star.beans import PropertyValue

from easydev.setting import DESKTOP, OS, WIN, WRITER, TOOLKIT


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

def msgbox(message, type_msg='infobox', title='Debug', buttons=1):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
    """
    desktop = _create_instance()
    toolkit = _create_instance(TOOLKIT, False)
    parent = toolkit.getDesktopWindow()
    mb = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return mb.execute()

def cmd(command, data):
    """
        Execute methods by name
    """
    return globals()[command](data)

def new_doc(type_doc='scalc'):
    """
        Create new doc
        scalc
        swriter
        simpress
        sdraw
        smath
    """
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
    """
    properties = _make_properties(options)
    path_url = path_to_url(path)
    desktop = _create_instance()
    doc = desktop.loadComponentFromURL(path_url, '_blank', 0, properties)
    return doc

def array(array, method, data):
    """
        Methods of list from Basic
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
        string identifying platform with as much useful information as possible,
        Returns the (real) processor name
    """
    info = (
        getpass.getuser(),
        platform.node(),
        platform.system(),
        platform.machine(),
        platform.platform(),
        platform.processor(),
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

def path_join(paths):
    return os.path.join(*paths)

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

