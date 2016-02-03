# coding: utf-8

import logging
from easydev.loapp import LOApp
from org.universolibre.EasyDev import XLOWriter
from easydev.setting import LOG, NAME_EXT, PY2
from easydev import comun

if PY2:
    str = unicode

log = logging.getLogger(NAME_EXT)


class LOWriter(XLOWriter, LOApp):

    def __init__(self, ctx, sm, desktop, toolkit):
        LOApp.__init__(self, ctx, sm, desktop, toolkit)

    def getParagraphs(self, doc, empty):
        paragraphs = []
        enum = doc.Text.createEnumeration()
        while enum.hasMoreElements():
            t = enum.nextElement()
            if t.supportsService('com.sun.star.text.Paragraph'):
                if empty:
                    paragraphs.append(t)
                else:
                    if t.String:
                        paragraphs.append(t)
        return tuple(paragraphs)

    def goStart(self, doc):
        cursor = doc.getCurrentController().getViewCursor()
        cursor.gotoStart(False)
        return cursor

    def goEnd(self, doc):
        cursor = doc.getCurrentController().getViewCursor()
        cursor.gotoEnd(False)
        return cursor

    def selectText(self, doc, text):
        doc.getCurrentController().select(text)
        return

