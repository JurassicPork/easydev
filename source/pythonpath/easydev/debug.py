# coding: utf-8

import sys
import logging
from datetime import datetime
from pprint import pprint
from org.universolibre.EasyDev import XDebug
from easydev import comun
from easydev.loapp import LOApp
from easydev.setting import (
    BUTTONS_OK,
    ERRORBOX,
    FILE_NAME_DEBUG,
    INFOBOX,
    NAME_EXT,
    OS,
    TITLE_DEBUG,
    WIN,
    WRITER,
)


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

    def __init__(self, ctx, sm, desktop, toolkit):
        self.ctx = ctx
        self.sm = sm
        self.desktop = desktop
        self.toolkit = toolkit

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
            app = LOApp(self.ctx, self.sm, self.desktop, self.toolkit)
            doc = app.getDoc(FILE_NAME_DEBUG)
            if not doc:
                doc = app.newDoc(WRITER)
            out = OutputDoc(doc)
            sys.stdout = out
        pprint (data)
        return

    def log(self, pathLog, data):
        path_log = comun.path_to_os(pathLog)
        with open(path_log, 'a') as out:
            out.write('{} - {} - '.format(str(datetime.now())[:19], NAME_EXT))
            pprint(data, stream=out)
        return

    def msgbox(self, message, type_msg=INFOBOX):
        """ Create message box
            type_msg: infobox, warningbox, errorbox, querybox, messbox
        """
        parent = self.toolkit.getDesktopWindow()
        mb = self.toolkit.createMessageBox(
            parent, type_msg, BUTTONS_OK, TITLE_DEBUG, str(message))
        return mb.execute()

    def mri(self, obj):
        m = self._create_instance('mytools.Mri')
        if m is None:
            self.msgbox('La extensión MRI no esta instalada', ERRORBOX)
            return
        m.inspect(obj)
        return

    def test(self, args):
        log.info(args)
        return True

