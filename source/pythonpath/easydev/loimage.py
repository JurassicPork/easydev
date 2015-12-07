# coding: utf-8

import logging
from org.universolibre.EasyDev import XLOImage
from easydev import comun
from easydev.comun import LODefault
from easydev.setting import LOG, NAME_EXT, PY2, SRV_GOS, SRV_GP, SHAPES, SRV_DCS

if PY2:
    str = unicode

log = logging.getLogger(NAME_EXT)


class LOImage(XLOImage, LODefault):

    def __init__(self, ctx, sm, desktop, toolkit):
        LODefault.__init__(self, ctx, sm, desktop, toolkit)

    def imageAdd(self, data):
        doc = data.Doc
        sheet = comun.get_sheet(doc, data.Sheet)
        properties = comun.to_dict(data.Properties)
        position, size = comun.get_pos_size(data.PosSize)
        draw_page = sheet.getDrawPage()

        if data.Type:
            type_shape = SHAPES.get(data.Type, SRV_DCS)
            image = doc.createInstance(type_shape)
            if type_shape == SRV_DCS:
                pv = comun.set_properties((('Type', data.Type),))
                comun.set_property(image, ('CustomShapeGeometry', pv))
        else:
            path = comun.path_to_url(data.Path)
            image = doc.createInstance(SRV_GOS)
            if data.Link:
                log.info(data.Link)
                image.GraphicURL = path
                log.info(path)
            else:
                gp = self._create_instance(SRV_GP)
                pv = comun.set_properties((('URL', path),))
                image.Graphic = gp.queryGraphic(pv)
        draw_page.add(image)

        if data.Name:
            image.Name = data.Name

        image.setPosition(position)
        image.setSize(size)
        for k, v in properties.items():
            if hasattr(image, k):
                setattr(image, k, v)
        return image

    def imageFormat(self, image, data):
        properties = comun.to_dict(data)
        title = None
        if image.supportsService('com.sun.star.chart2.Title'):
            title = image.getText()[0]
        elif image.supportsService('com.sun.star.chart.ChartAxis'):
            title = image.getAxisTitle()
        for k, v in properties.items():
            if hasattr(image, k):
                setattr(image, k, v)
            if title and hasattr(title, k):
                setattr(title, k, v)
        return
