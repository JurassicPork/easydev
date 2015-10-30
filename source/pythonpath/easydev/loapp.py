# coding: utf-8

import uno
import logging
from org.universolibre.EasyDev import XLOApp
from easydev import comun
from easydev.comun import LODefault
from easydev.setting import LOG, NAME_EXT, CALC, EXT_PDF


log = logging.getLogger(NAME_EXT)


class LOApp(XLOApp, LODefault):

    def __init__(self, ctx, sm, desktop, toolkit):
        LODefault.__init__(self, ctx, sm, desktop, toolkit)

    def newDoc(self, type_doc=CALC):
        """
            Create new doc
            http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XComponentLoader.html

        type_doc:
            scal
            swriter
            simpress
            sdraw
            smath
        """
        if not type_doc:
            type_doc = 'scalc'
        path = 'private:factory/{}'.format(type_doc)
        doc = self.desktop.loadComponentFromURL(path, '_default', 0, ())
        return doc

    def newDB(self, path):
        dbc = self._create_instance('com.sun.star.sdb.DatabaseContext')
        db = dbc.createInstance()
        db.URL = 'sdbc:embedded:hsqldb'
        db.DatabaseDocument.storeAsURL(comun.path_to_url(path), ())
        return db

    def getDoc(self, title):
        """
            If title is missing get current component,
            else search doc title in components
        """
        if not title:
            return self.desktop.getCurrentComponent()

        enum = self.desktop.getComponents().createEnumeration()
        while enum.hasMoreElements():
            doc = enum.nextElement()
            if doc.Title == title:
                return doc
        return None

    def getTypeDoc(self, doc):
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
        return ''

    def getDocs(self):
        """
            Return all documents open
        """
        docs = []
        enum = self.desktop.getComponents().createEnumeration()
        while enum.hasMoreElements():
            docs.append(enum.nextElement())
        return tuple(docs)

    def openDoc(self, path, options):
        """
            Open doc
            http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html
            http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1MediaDescriptor.html
        """
        properties = comun.set_properties(options)
        path_url = comun.path_to_url(path)
        doc = self.desktop.loadComponentFromURL(path_url, '_blank', 0, properties)
        return doc

    def setFocus(self, doc):
        """
            Active doc
        """
        window = doc.getCurrentController().getFrame().getComponentWindow()
        window.setFocus()
        return

    def getStatusBar(self, doc):
        """
            Return status bar
        """
        statusbar = doc.getCurrentController().getStatusIndicator()
        return statusbar

    def exportPDF(self, doc, path_save, options):
        """
            Export to PDF
            http://wiki.services.openoffice.org/wiki/API/Tutorials/PDF_export
        """
        close = False
        if isinstance(doc, str):
            close = True
            doc = self.openDoc(doc, (('Hidden', True),))
        if path_save:
            if comun.isdir(path_save):
                path_save = comun.replace_name_ext(path_save, doc.getURL(), EXT_PDF)
        else:
            path_save = comun.replace_ext(doc.getURL(), EXT_PDF)
        path_save = comun.path_to_url(path_save)
        try:
            filter_name = '{}_pdf_Export'.format(self.getTypeDoc(doc))
            if options:
                filter_data = comun.set_properties(options)
                filter_options = (
                    ('FilterName', filter_name),
                    ('FilterData', filter_data),
                )
            else:
                filter_options = (('FilterName', filter_name),)
            media_descriptor = comun.set_properties(filter_options)
            doc.storeToURL(path_save, media_descriptor)
        except:
            log.error('PDF', exc_info=True)
        if close:
            doc.dispose()
        if comun.exists(path_save):
            return comun.path_to_os(path_save)
        return ''

