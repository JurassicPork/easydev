# coding: utf-8

import unohelper
from com.sun.star.lang import XServiceName
from com.sun.star.lang import XServiceInfo
from easydev.debug import Debug
from easydev.tools import Tools
#~ from easydev.email import EMail
from easydev.localc import LOCalc
from easydev.setting import ID_EXT, SRV_JOB, DESKTOP, TOOLKIT


class EasyDev(unohelper.Base, XServiceName, XServiceInfo, Debug, Tools, LOCalc):

    service_name = implementation_name = ID_EXT

    def __init__(self, ctx=None):
        self.ctx = ctx
        self.sm = self.ctx.getServiceManager()
        self.desktop = self.sm.createInstanceWithContext(DESKTOP, self.ctx)
        self.toolkit = self.sm.createInstance(TOOLKIT)
        Debug.__init__(self, self.ctx, self.sm, self.desktop, self.toolkit)
        Tools.__init__(self, self.ctx, self.sm, self.desktop, self.toolkit)
        LOCalc.__init__(self, self.ctx, self.sm, self.desktop, self.toolkit)
        #~ EMail.__init__(self)

    def getServiceName(self):
        return self.service_name

    def getImplementationName(self):
        return self.implementation_name

    def getSupportedServiceNames(self):
        return (self.service_name,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(EasyDev, ID_EXT, SRV_JOB)
