# coding: utf-8

import logging
from easydev.loapp import LOApp
from org.universolibre.EasyDev import XLOCalc
from easydev.setting import LOG, NAME_EXT


log = logging.getLogger(NAME_EXT)


class LOCalc(XLOCalc, LOApp):

    def __init__(self, ctx, sm, desktop, toolkit):
        LOApp.__init__(self, ctx, sm, desktop, toolkit)

    def getCell(self, doc, sheet_name=None, cell_address=None):
        if isinstance(doc, str):
            doc = self.getDoc(doc)
        if not sheet_name and not cell_address:
            cell = doc.getCurrentSelection()
        else:
            if isinstance(sheet_name, str):
                if sheet_name:
                    sheet = doc.getSheets().getByName(sheet_name)
                else:
                    sheet = doc.getCurrentController().getActiveSheet()
            else:
                sheet = sheet_name
            if isinstance(cell_address, str):
                cell = sheet.getCellRangeByName(cell_address)
            else:
                cell = sheet.getCellByPosition(*cell_address)
        return cell
