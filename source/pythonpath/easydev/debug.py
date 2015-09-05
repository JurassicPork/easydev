# coding: utf-8

import logging
from pprint import pprint
from org.universolibre.util.EasyDev import XDebug
from easydev.setting import LOG, NAME_EXT, OS, WIN, BUTTONS_OK, DESKTOP, \
    TOOLKIT, WRITER


log = logging.getLogger(NAME_EXT)


class OutputDoc(object):

    def __init__(self, doc):
        self.doc = doc

    def write(self, info):
        text = self.doc.Text
        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        text.insertString(cursor, str(info), 0)
        return


class Debug(XDebug):
    error = ''

    def __init__(self, ctx, sm):
        self.ctx = ctx
        self.sm = sm

    def _create_instance(self, name, with_context=True):
        if with_context:
            instance = self.sm.createInstanceWithContext(name, self.ctx)
        else:
            instance = self.sm.createInstance(name)
        return instance

    def debug(self, data):
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

    def msgbox(self, message, type_msg='infobox'):
        """ Create message box
            type_msg: infobox, warningbox, errorbox, querybox, messbox
        """
        desktop = self._create_instance(DESKTOP)
        toolkit = self._create_instance(TOOLKIT, False)
        parent = toolkit.getDesktopWindow()
        mb = toolkit.createMessageBox(
            parent, 'infobox', BUTTONS_OK, 'Debug', str(message))
        return mb.execute()

    def mri(self, obj):
        m = self._create_instance('mytools.Mri')
        if m is None:
            self.msgbox('La extensión MRI no esta instalada', 'errorbox')
            return
        m.inspect(obj)
        return
