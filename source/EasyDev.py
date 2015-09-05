# coding: utf-8

import unohelper
from com.sun.star.lang import XServiceName
from com.sun.star.lang import XServiceInfo
from easydev.debug import Debug
from easydev.tools import Tools
from easydev.email import EMail
#~ from easydev.loapp import LOApp
from easydev.localc import LOCalc
from easydev.setting import ID_EXT, SRV_JOB


class EasyDev(unohelper.Base, XServiceName, XServiceInfo, Debug, Tools, EMail,
    LOCalc):

    service_name = implementation_name = ID_EXT

    def __init__(self, ctx=None):
        self.ctx = ctx
        self.sm = self.ctx.getServiceManager()
        Debug.__init__(self, self.ctx, self.sm)
        Tools.__init__(self, self.ctx, self.sm)
        EMail.__init__(self)
        #~ LOApp.__init__(self, self.ctx, self.sm)
        LOCalc.__init__(self, self.ctx, self.sm)

    def getTypeDoc(self, doc):
        return tools.get_type_doc(doc)

    def getDocs(self):
        return tools.get_docs()

    def openDoc(self, pathDoc, options):
        return tools.open_doc(pathDoc, options)

    def setFocus(self, doc):
        return tools.set_focus(doc)

    def getStatusBar(self, doc):
        return tools.get_status_bar(doc)

    def exportPDF(self, doc, pathSave, options):
        return tools.export_pdf(doc, pathSave, options)

    def getPath(self, name):
        return tools.get_path(name)

    def getPathInfo(self, path):
        return tools.get_path_info(path)

    def pathJoin(self, paths):
        return tools.path_join(paths)

    def getFolder(self, initFolder):
        return tools.get_folder(initFolder)

    def getSelectedFiles(self, initFolder, multiple, filters):
        return tools.get_selected_files(initFolder, multiple, filters)

    def getFiles(self, path, ext):
        return tools.get_files(path, ext)

    def execute(self, args, wait):
        return tools.execute(args, wait)

    def getConfig(self, key):
        return tools.get_config(key)

    def setConfig(self, key, value):
        return tools.set_config(key, value)

    def render(self, template, data):
        return tools.render(template, data)

    def fileOpen(self, path, mode):
        return tools.file_open(path, mode)

    def format(self, template, data):
        return tools.format(template, data)

    def question(self, title, message):
        return tools.question(title, message)

    def getClipboard(self):
        return tools.get_text_from_clipboard()

    def setClipboard(self, text):
        return tools.set_text_to_clipboard(text)

    def exportCSV(self, path, data, options):
        return tools.export_csv(path, data, options)

    def getRange(self, doc, sheetName, rangeAddress):
        return tools.get_range(doc, sheetName, rangeAddress)

    def selectRange(self, doc, sheetName, rango):
        return tools.select_range(doc, sheetName, rango)

    def getCurrentRegion(self, cell):
        return tools.get_current_region(cell)

    def getLastRow(self, cell):
        return tools.get_last_row(cell)

    def getServiceName(self):
        return self.service_name

    def getImplementationName(self):
        return self.implementation_name

    def getSupportedServiceNames(self):
        return (self.service_name,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(EasyDev, ID_EXT, SRV_JOB)
