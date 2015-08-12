# coding: utf-8
import uno
import ctypes
import subprocess
import sys
import getpass
import platform
from pprint import pprint
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

