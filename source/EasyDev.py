# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.lang import XServiceName
from com.sun.star.lang import XServiceInfo
from org.universolibre.util.EasyDev import XEasyDev
from easydev import tools
from easydev.setting import VERSION, ID_EXT, NAME_EXT, CALC


SRV_JOB = ('com.sun.star.task.Job',)


class EasyDev(unohelper.Base, XServiceName, XServiceInfo, XEasyDev):

    version = VERSION
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

    def getDocs(self):
        return tools.get_docs()

    def openDoc(self, pathDoc, options):
        return tools.open_doc(pathDoc, options)

    def array(self, array, method, data):
        return tools.array(array, method, data)

    def getSizeScreen(self):
        return tools.get_size_screen()

    def getInfoPC(self):
        return tools.get_info_pc()

    def getServiceName(self):
        return self.service_name

    def getImplementationName(self):
        return self.implementation_name

    def getSupportedServiceNames(self):
        return (self.service_name,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(EasyDev, ID_EXT, SRV_JOB)
