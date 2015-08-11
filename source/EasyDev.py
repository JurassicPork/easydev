# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.lang import XServiceName
from com.sun.star.lang import XServiceInfo
from org.universolibre.util.EasyDev import XEasyDev
from easydev import util


ID = 'org.universolibre.util.EasyDev'
SERVICE = ('com.sun.star.task.Job',)


class EasyDev(unohelper.Base, XServiceName, XServiceInfo, XEasyDev):

    version = '0.1.0'
    service_name = implementation_name = "EasyDev"

    def __init__(self, ctx=None):
        self.ctx = ctx

    def debug(self, data):
        return util.debug(data)

    def msgbox(self, data):
        return util.msgbox(data)

    def getSizeScreen(self):
        return util.get_size_screen()

    def getDoc(self, title=''):
        return util.get_doc(title)

    def newDoc(self, typedoc=0):
        return util.new_doc(typedoc)

    def getInfoPC(self):
        return util.get_info_pc()

    def getServiceName(self):
        return self.service_name

    def getImplementationName(self):
        return self.implementation_name

    def getSupportedServiceNames(self):
        return (self.service_name,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(EasyDev, ID, SERVICE)
