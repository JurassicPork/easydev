# coding: utf-8

import logging
from easydev import comun
from easydev.loapp import LOApp
from org.universolibre.EasyDev import XLOCalc
from easydev.setting import LOG, NAME_EXT


log = logging.getLogger(NAME_EXT)


class LOCalc(XLOCalc, LOApp):

    def __init__(self, ctx, sm, desktop, toolkit):
        LOApp.__init__(self, ctx, sm, desktop, toolkit)

    def getCell(self, address):
        if not address.Doc or isinstance(address.Doc, str):
            doc = self.getDoc(address.Doc)
        else:
            doc = address.Doc
        if address.Current:
            return doc.getCurrentSelection().getCellByPosition(0, 0)
        if not address.Sheet:
            sheet = doc.getCurrentController().getActiveSheet()
        elif isinstance(address.Sheet, str):
            sheet = doc.getSheets().getByName(address.Sheet)
        else:
            sheet = address.Sheet
        if address.Name:
            cell = sheet.getCellRangeByName(address.Name.upper()).getCellByPosition(0, 0)
        else:
            cell = sheet.getCellByPosition(address.Col, address.Row)
        return cell

    def getRange(self, address):
        if not address.Doc or isinstance(address.Doc, str):
            doc = self.getDoc(address.Doc)
        else:
            doc = address.Doc
        if address.Current:
            return doc.getCurrentSelection()
        if not address.Sheet:
            sheet = doc.getCurrentController().getActiveSheet()
        elif isinstance(address.Sheet, str):
            sheet = doc.getSheets().getByName(address.Sheet)
        else:
            sheet = address.Sheet
        if address.Name:
            rango = sheet.getCellRangeByName(address.Name.upper())
        else:
            rango = sheet.getCellRangeByPosition(
                address.Col, address.Row, address.EndCol, address.EndRow)
        return rango

    def selectRange(self, doc, rango):
        doc.getCurrentController().select(rango)
        return

    def getCurrentRegion(self, cell, getCursor):
        sheet = cell.getSpreadsheet()
        cursor = sheet.createCursorByRange(cell)
        cursor.collapseToCurrentRegion()
        if getCursor:
            return cursor
        else:
            return sheet.getCellRangeByName(cursor.AbsoluteName)

    def getLastRow(self, cell):
        cursor = cell.getSpreadsheet().createCursorByRange(cell)
        cursor.gotoEnd()
        return cursor.getRangeAddress().EndRow

    def getNextID(self, cell, col):
        cr = self.getCurrentRegion(cell, False)
        data = cr.getDataArray()
        column = [row[col] for row in data if isinstance(row[col], (int, float))]
        value = int(max(column))
        return value + 1

    def getValue(self, cell):
        get_type = cell.getType()
        if not get_type:
            return ''
        if get_type == 1:
            return cell.getValue()
        if get_type == 2:
            return cell.getString()
        return cell.getFormula()

    def setValue(self, cell, value):
        if isinstance(value, str):
            if value.startswith('='):
                cell.setFormula(value)
            else:
                cell.setString(value)
        elif isinstance(value, (int, float)):
            cell.setValue(value)
        return

    def setData(self, cell, data):
        rango = comun.offset(cell, len(data), len(data[0]), True)
        rango.setDataArray(data)
        return

    def getEmpty(self, cell):
        rango = cell
        if comun.is_cell(cell):
            rango = self.getCurrentRegion(cell, False)
        return rango.queryEmptyCells()

