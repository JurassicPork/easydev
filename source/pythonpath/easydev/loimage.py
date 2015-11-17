# coding: utf-8

import logging
from org.universolibre.EasyDev import XLOImage
from easydev import comun
from easydev.comun import LODefault
from easydev.setting import LOG, NAME_EXT, SRV_GOS, SRV_GP, SHAPES, SRV_DCS


log = logging.getLogger(NAME_EXT)


class LOImage(XLOImage, LODefault):

    def __init__(self, ctx, sm, desktop, toolkit):
        LODefault.__init__(self, ctx, sm, desktop, toolkit)

    def imageAdd(self, data):
        doc = data.Doc
        sheet = comun.get_sheet(doc, data.Sheet)
        path = comun.path_to_url(data.Path)
        draw_page = sheet.getDrawPage()
        if data.Type:
            type_shape = SHAPES.get(data.Type, SRV_DCS)
            image = doc.createInstance(type_shape)
            if type_shape == SRV_DCS:
                pv = comun.set_properties((('Type', data.Type),))
                image.setPropertyValue('CustomShapeGeometry', pv)
        else:
            image = doc.createInstance(SRV_GOS)
            if data.Link:
                image.GraphicURL = path
            else:
                gp = self._create_instance(SRV_GP)
                pv = comun.set_properties((('URL', path),))
                image.Graphic = gp.queryGraphic(pv)
        draw_page.add(image)

        if data.Name:
            image.Name = data.Name

        position, size = comun.get_pos_size(data.PosSize)
        image.setPosition(position)
        image.setSize(size)
        return image

