# coding: utf-8

import logging
from org.universolibre.EasyDev import XLOApp
from easydev.setting import LOG, NAME_EXT, CALC


log = logging.getLogger(NAME_EXT)


class LOApp(XLOApp):

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

    def getDoc(self, title=''):
        """
            If title is missing get current component,
            else search doc title in components
        """
        if not title:
            return self.desktop.getCurrentComponent()

        enum = self.desktop.getComponents().createEnumeration()
        while enum.hasMoreElements():
            doc = enum.nextElement()
            if doc.getTitle() == title:
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
        properties = make_properties(options)
        path_url = path_to_url(path)
        doc = self.desktop.loadComponentFromURL(path_url, '_blank', 0, properties)
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
            _, _, name, extension = get_path_info(path_to_os(doc.getURL()))
            path_save = path_to_url(
                os.path.normpath('{}/{}.{}'.format(path_save, name, EXT_PDF)))
        else:
            path_save = path_to_url(path_save)
    else:
        path_save = path_to_url(replace_ext(path_to_os(doc.getURL()), EXT_PDF))
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
    if os.path.exists(path_to_os(path_save)):
        return path_save
    return ''


def get_range(doc, sheet_name=None, range_address=None):
    if isinstance(doc, str):
        doc = get_doc(doc)
    if not sheet_name and not range_address:
        rango = doc.getCurrentSelection()
    else:
        if isinstance(sheet_name, str):
            if sheet_name:
                sheet = doc.getSheets().getByName(sheet_name)
            else:
                sheet = doc.getCurrentController().getActiveSheet()
        else:
            sheet = sheet_name
        if isinstance(range_address, str):
            rango = sheet.getCellRangeByName(range_address)
        else:
            rango = sheet.getCellRangeByPosition(*range_address)
    return rango


def select_range(doc, sheet_name=None, rango=None):
    if isinstance(doc, str):
        doc = get_doc(doc)
    if isinstance(sheet_name, str):
        if sheet_name:
            sheet = doc.getSheets().getByName(sheet_name)
        else:
            sheet = doc.getCurrentController().getActiveSheet()
    else:
        sheet = sheet_name
    if isinstance(rango, str):
        rango = sheet.getCellRangeByName(rango)
    doc.getCurrentController().select(rango)
    return


def get_current_region(cell):
    cursor = cell.getSpreadsheet().createCursorByRange(cell)
    cursor.collapseToCurrentRegion()
    return cursor


def get_last_row(cell):
    cursor = cell.getSpreadsheet().createCursorByRange(cell)
    cursor.gotoEnd()
    return cursor.getRangeAddress().EndRow

