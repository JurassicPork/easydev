# coding: utf-8

import logging
from org.universolibre.EasyDev import XLOApp
from easydev.setting import LOG, NAME_EXT, CALC, DESKTOP


log = logging.getLogger(NAME_EXT)


class LOApp(XLOApp):

    def __init__(self, ctx, sm):
        self.ctx = ctx
        self.sm = sm

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
        desktop = self._create_instance(DESKTOP)
        path = 'private:factory/{}'.format(type_doc)
        doc = desktop.loadComponentFromURL(path, '_default', 0, ())
        return doc

    def getDoc(self, title=''):
        """
            If title is missing get current component,
            else search doc title in components
        """
        desktop = self._create_instance(DESKTOP)
        if not title:
            return desktop.getCurrentComponent()

        enum = desktop.getComponents().createEnumeration()
        while enum.hasMoreElements():
            doc = enum.nextElement()
            if doc.getTitle() == title:
                return doc
        return None
