# coding: utf-8

import logging
from easydev.loapp import LOApp
from org.universolibre.EasyDev import XLOWriter
from easydev.setting import LOG, NAME_EXT


log = logging.getLogger(NAME_EXT)


class LOWriter(XLOWriter, LOApp):

    def __init__(self, ctx, sm, desktop, toolkit):
        LOApp.__init__(self, ctx, sm, desktop, toolkit)

    def getParagraphs(self, doc):
        paragraphs = []
        enum = doc.Text.createEnumeration()
        while enum.hasMoreElements():
            t = enum.nextElement()
            if t.supportsService('com.sun.star.text.Paragraph'):
                paragraphs.append(t)
        return tuple(paragraphs)
