# coding: utf-8
import uno
import ctypes
import subprocess
import sys
import getpass
import platform
from pprint import pprint
from easydev.setting import OS, WIN

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


def _create_instance(name, with_context=False):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance

def debug(data):
    if OS == WIN:
        doc = get_doc('debug.odt')
        if not doc:
            doc = new_doc(1)
        out = OutputDoc(doc)
        sys.stdout = out
    pprint (data)
    return

def msgbox(message, type_msg='infobox', title='Debug', buttons=1):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
    """
    desktop = _create_instance('com.sun.star.frame.Desktop', True)
    toolkit = _create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    mb = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return mb.execute()

def get_size_screen():
    if OS == WIN:
        user32 = ctypes.windll.user32
        res = '{}x{}'.format(user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
    else:
        args = 'xrandr | grep "\*" | cut -d" " -f4'
        res = subprocess.check_output(args, shell=True).decode()
        return res

def get_doc(title=''):
    """
        If title is missing get current component,
        else search doc title in components
    """
    desktop = _create_instance('com.sun.star.frame.Desktop', True)
    if not title:
        return desktop.getCurrentComponent()

    enum = desktop.getComponents().createEnumeration()
    while enum.hasMoreElements():
        doc = enum.nextElement()
        if doc.getTitle() == title:
            return doc
    return None

def new_doc(typedoc=0):
    """
        Create new doc
        scalc = 0
        swriter = 1
        simpress = 2
        sdraw = 3
        smath = 4
    """
    desktop = _create_instance('com.sun.star.frame.Desktop', True)
    types = ['scalc', 'swriter', 'simpress', 'sdraw', 'smath']
    path = 'private:factory/{}'.format(types[typedoc])
    doc = desktop.loadComponentFromURL(path, '_default', 0, ())
    return doc

def get_info_pc():
    info = (
        getpass.getuser(),
        platform.node(),
        platform.system(),
        platform.machine(),
        platform.platform(),
        platform.processor(),
    )
    return info

