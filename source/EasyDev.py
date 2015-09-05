# coding: utf-8

import unohelper
from com.sun.star.lang import XServiceName
from com.sun.star.lang import XServiceInfo
from easydev.debug import Debug
from easydev.tools import Tools
from easydev.email import EMail
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
        LOCalc.__init__(self, self.ctx, self.sm)

    def getServiceName(self):
        return self.service_name

    def getImplementationName(self):
        return self.implementation_name

    def getSupportedServiceNames(self):
        return (self.service_name,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(EasyDev, ID_EXT, SRV_JOB)
