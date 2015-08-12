# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.lang import XServiceName
from com.sun.star.lang import XServiceInfo
from org.universolibre.util.EasyDev import XEasyDev
from easydev import tools


ID = 'org.universolibre.util.EasyDev'
SERVICE = ('com.sun.star.task.Job',)


class EasyDev(unohelper.Base, XServiceName, XServiceInfo, XEasyDev):

    version = '0.1.0'
    service_name = implementation_name = "EasyDev"

    def __init__(self, ctx=None):
        self.ctx = ctx

    def debug(self, data):
        return tools.debug(data)

    def msgbox(self, data):
        return tools.msgbox(data)

    def newDoc(self, typeDoc='scalc'):
        return tools.new_doc(typeDoc)

    def getDoc(self, title=''):
        return tools.get_doc(title)

    def getDocs(self):
        return tools.get_docs()

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
g_ImplementationHelper.addImplementation(EasyDev, ID, SERVICE)
