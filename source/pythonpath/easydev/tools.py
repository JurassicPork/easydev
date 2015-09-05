# coding: utf-8

import sys
import os
import re
import json
import logging
import ctypes
import subprocess
import getpass
import platform
import csv
from string import Template

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.datatransfer import XTransferable, DataFlavor
from org.universolibre.EasyDev import XTools

from easydev.setting import LOG, NAME_EXT, OS, VERSION, DESKTOP, TOOLKIT,
    BUTTONS_YES_NO, YES, NODE, NODE_CONFIG, CLIPBOARD_FORMAT_TEXT


log = logging.getLogger(NAME_EXT)


class Tools(XTools):
    VERSION = VERSION
    OS = OS
    value = ''

    def __init__(self, ctx, sm):
        self.ctx = ctx
        self.sm = sm

    def _create_instance(self, name, with_context=True):
        if with_context:
            instance = self.sm.createInstanceWithContext(name, self.ctx)
        else:
            instance = self.sm.createInstance(name)
        return instance

    def getSizeScreen(self):
        if OS == WIN:
            user32 = ctypes.windll.user32
            res = '{}x{}'.format(
                user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
        else:
            args = 'xrandr | grep "\*" | cut -d" " -f4'
            res = subprocess.check_output(args, shell=True).decode()
            return res

    def getInfoPC(self):
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

    def question(self, title, message):
        """ Create message box
            type_msg: infobox, warningbox, errorbox, querybox, messbox
        """
        desktop = self._create_instance(DESKTOP)
        toolkit = self._create_instance(TOOLKIT, False)
        parent = toolkit.getDesktopWindow()
        mb = toolkit.createMessageBox(
            parent, 'querybox', BUTTONS_YES_NO, title, message)
        return YES == mb.execute()

    def render(self, template, data):
        data = {r[0]: r[1] for r in data}
        s = Template(template)
        return s.safe_substitute(**data)

    def format(self, template, data):
        if isinstance(data, tuple):
            if isinstance(data[0], tuple):
                data = {r[0]: r[1] for r in data}
                result = template.format(**data)
            else:
                result = template.format(*data)
        else:
            result = template.format(data)
        return result

    def path_to_os(self, path):
        if path.startswith('file://'):
            path = uno.fileUrlToSystemPath(path)
        return path

    def path_to_url(self, path):
        if path.startswith('file://'):
            return path
        return uno.systemPathToFileUrl(path)

    def getPath(self, name):
        """
            Return de path name in config
            http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XPathSettings.html
        """
        if not name:
            name = 'Work'
        path = self._create_instance('com.sun.star.util.PathSettings')
        return self.path_to_os(getattr(path, name))

    def getPathInfo(self, path):
        path = self.path_to_os(path)
        path, filename = os.path.split(path)
        name, extension = os.path.splitext(filename)
        return (path, filename, name, extension)

    def pathJoin(self, paths):
        return os.path.normpath(os.path.join(*paths))

    def getFolder(self, init_folder=''):
        if init_folder:
            init_folder = self.path_to_url(init_folder)
        else:
            init_folder = self.getPath('Work')
        folder = self._create_instance('com.sun.star.ui.dialogs.FolderPicker')
        folder.setDisplayDirectory(init_folder)
        if folder.execute():
            return folder.getDirectory()
        else:
            return ''

    def getSelectedFiles(self, init_folder, multiple, filters):
        if init_folder:
            init_folder = self.path_to_url(init_folder)
        else:
            init_folder = self.getPath('Work')

        folder = self._create_instance('com.sun.star.ui.dialogs.FilePicker')
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

    def getFiles(self, path, ext):
        paths = []
        for folder, _, files in os.walk(path):
            pattern = re.compile('\.{}'.format(ext), re.IGNORECASE)
            paths += [os.path.join(folder, f) for f in files if pattern.search(f)]
        return tuple(paths)

    def fileOpen(self, path, mode='r'):
        data = ''
        if not mode:
            mode = 'r'
        path = path_to_os(path)
        with open(path, mode) as f:
            data = f.read()
        return data

    def execute(self, args, wait):
        if wait:
            res = subprocess.check_output(' '.join(args), shell=True).decode()
            return res
        else:
            subprocess.Popen(args)
        return

    def getConfig(self, key):
        name = 'com.sun.star.configuration.ConfigurationProvider'
        cp = self._create_instance(name)
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
            log.debug(e)
            return

    def setConfig(self, key, value):
        name = 'com.sun.star.configuration.ConfigurationProvider'
        cp = _create_instance(name)
        node = PropertyValue()
        node.Name = 'nodepath'
        node.Value = NODE
        try:
            config_writer = cp.createInstanceWithArguments(
                'com.sun.star.configuration.ConfigurationUpdateAccess', (node,))
            data = self.getConfig('')
            if not data:
                data = {}
            data[key] = value
            new_values = json.dumps(data)
            config_writer.setPropertyValue(NODE_CONFIG, new_values)
            config_writer.commitChanges()
            return True
        except Exception as e:
            log.debug(e)
            return False

    def getClipboard(self):
        df = None
        text = ''
        sc = self._create_instance(
            'com.sun.star.datatransfer.clipboard.SystemClipboard')
        transferable = sc.getContents()
        data = transferable.getTransferDataFlavors()
        for df in data:
            if df.MimeType == CLIPBOARD_FORMAT_TEXT:
                break
        if df:
            text = transferable.getTransferData(df)
        return text

    def setClipboard(self, text):
        ts = TextTransferable(text)
        sc = self._create_instance(
            'com.sun.star.datatransfer.clipboard.SystemClipboard')
        sc.setContents(ts, None)
        return

    def array(self, array, method, data):
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


class TextTransferable(unohelper.Base, XTransferable):
    """Keep clipboard data and provide them."""

    def __init__(self, text):
        df = DataFlavor()
        df.MimeType = CLIPBOARD_FORMAT_TEXT
        df.HumanPresentableName = "encoded text utf-16"
        self.flavors = [df]
        self.data = [text]

    def getTransferData(self, flavor):
        if not flavor:
            return
        for i, f in enumerate(self.flavors):
            if flavor.MimeType == f.MimeType:
                return self.data[i]
        return

    def getTransferDataFlavors(self):
        return tuple(self.flavors)

    def isDataFlavorSupported(self, flavor):
        if not flavor:
            return False
        mtype = flavor.MimeType
        for f in self.flavors:
            if mtype == f.MimeType:
                return True
        return False


def _make_properties(properties):
    prop = []
    l = len(properties)
    for i in range(0, l, 2):
        pv = PropertyValue()
        pv.Name = properties[i]
        pv.Value = properties[i + 1]
        prop.append(pv)
    return tuple(prop)


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











def replace_ext(path, ext):
    path, _, name, _ = get_path_info(path)
    return '{}/{}.{}'.format(path, name, ext)


















def export_csv(path, data, options=()):
    try:
        headers = ()
        config = {}
        path = path_to_os(path)
        if options:
            config = {r[0]: r[1] for r in options}
        write_headers = config.get('write_headers', False)
        config.pop('write_headers', None)
        if write_headers:
            headers = config.get('headers', ())
            config.pop('headers', None)
            if not headers:
                headers = data[0]
            data = data[1:]
        with open(path, 'w') as f:
            if config:
                writer = csv.writer(f, **config)
            else:
                writer = csv.writer(f)
            if headers:
                writer.writerow(headers)
            writer.writerows(data)
        return True
    except:
        log.debug('CSV', exc_info=True)
        return False


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

