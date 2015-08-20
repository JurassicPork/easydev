# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.lang import XServiceName
from com.sun.star.lang import XServiceInfo
from org.universolibre.util.EasyDev import XEasyDev
from easydev import tools
from easydev.setting import VERSION, ID_EXT, NAME_EXT, CALC, OS


SRV_JOB = ('com.sun.star.task.Job',)


class EasyDev(unohelper.Base, XServiceName, XServiceInfo, XEasyDev):

    VERSION = VERSION
    OS = OS
    value = ''
    service_name = implementation_name = NAME_EXT

    def __init__(self, ctx=None):
        self.ctx = ctx

    def debug(self, data):
        return tools.debug(data)

    def msgbox(self, data):
        return tools.msgbox(data)

    def cmd(self, command, data):
        return tools.cmd(command, data)

    def newDoc(self, typeDoc=CALC):
        return tools.new_doc(typeDoc)

    def getDoc(self, title=''):
        return tools.get_doc(title)

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

    def array(self, array, method, data):
        return tools.array(array, method, data)

    def getSizeScreen(self):
        return tools.get_size_screen()

    def getInfoPC(self):
        return tools.get_info_pc()

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

    def sendMail(self, server, mail, files):
        return tools.send_mail(server, mail, files)

    def render(self, template, data):
        return tools.render(template, data)

    def fileOpen(self, path, mode):
        return tools.file_open(path, mode)

    def format(self, template, data):
        return tools.format(template, data)

    def question(self, title, message):
        return tools.question(title, message)

    def getServiceName(self):
        return self.service_name

    def getImplementationName(self):
        return self.implementation_name

    def getSupportedServiceNames(self):
        return (self.service_name,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(EasyDev, ID_EXT, SRV_JOB)
