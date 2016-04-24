# coding: utf-8

import logging
from easydev import comun
from easydev.loapp import LOApp
from org.universolibre.EasyDev import XLOCalc
from easydev.setting import LOG, NAME_EXT, PY2, SRV_GOS


if PY2:
    str = unicode

log = logging.getLogger(NAME_EXT)


class LOCalc(XLOCalc, LOApp):

    def __init__(self, ctx, sm, desktop, toolkit):
        LOApp.__init__(self, ctx, sm, desktop, toolkit)

    def _get_doc(self, document):
        if not document or isinstance(document, str):
            document = self.getDoc(document)
        return document

    def _get_new_name(self, doc, name):
        new_name = name
        i = 1
        while doc.getSheets().hasByName(new_name):
            new_name = '{}_{}'.format(name, i)
            i += 1
        return new_name

    def getSheetsNames(self, doc):
        return doc.getSheets().getElementNames()

    def getSheet(self, address):
        doc = self._get_doc(address.Doc)
        return comun.get_sheet(doc, address.Sheet)

    def sheetActivate(self, address):
        doc = self._get_doc(address.Doc)
        address.Doc = doc
        sheet = self.getSheet(address)
        doc.getCurrentController().setActiveSheet(sheet)
        return sheet

    def sheetInsert(self, address, pos, rename):
        doc = self._get_doc(address.Doc)
        index = pos
        if pos < 0:
            index = doc.getSheets().getCount() + pos + 1
        names = address.Sheet
        if isinstance(names, str):
            if rename:
                names = self._get_new_name(doc, names)
            doc.getSheets().insertNewByName(names, index)
        elif isinstance(names, tuple):
            for name in names:
                if rename:
                    name = self._get_new_name(doc, name)
                doc.getSheets().insertNewByName(name, index)
        return

    def sheetInsertFromDoc(self, source, target, pos, values):
        src_doc = self._get_doc(source.Doc)
        doc = target.Doc
        index = pos
        if pos < 0:
            index = doc.getSheets().getCount() + pos + 1
        if source.Sheet:
            names = self._get_sheets_names(source.Sheet, src_doc, False)
        else:
            names = self.getSheetsNames(src_doc)
        sheet = None
        for name in names:
            new_pos = doc.getSheets().importSheet(src_doc, name, index)
            sheet = doc.getSheets().getByIndex(new_pos)
            if values:
                pass
        doc.getCurrentController().setActiveSheet(sheet)
        doc.getCurrentController().setFormDesignMode(False)
        return sheet

    def _get_sheet_name(self, name, doc):
        if isinstance(name, int):
            index = name
            if index < 0:
                index = doc.getSheets().getCount() + index
            name = doc.getSheets().getByIndex(index).getName()
        elif not isinstance(name, str):
            name = name.getName()
        return name

    def _get_sheets_names(self, name, doc, allSheets):
        if isinstance(name, tuple):
            names = tuple(self._get_sheet_name(n, doc) for n in name)
        else:
            names = (self._get_sheet_name(name, doc),)
        if allSheets:
            all_names = self.getSheetsNames(doc)
            names = set(names).symmetric_difference(all_names)
        return names

    def sheetRemove(self, address, allSheets):
        doc = self._get_doc(address.Doc)
        names = self._get_sheets_names(address.Sheet, doc, allSheets)
        for n in names:
            doc.getSheets().removeByName(n)
        return

    def sheetMove(self, address, pos):
        doc = self._get_doc(address.Doc)
        names = self._get_sheets_names(address.Sheet, doc, False)
        index = pos
        if pos < 0:
            index = doc.getSheets().getCount() + pos + 1
        for n in names:
            doc.getSheets().moveByName(n, index)
        return

    def sheetName(self, address, newname):
        doc = self._get_doc(address.Doc)
        names = self._get_sheets_names(address.Sheet, doc, False)
        for n in names:
            new_name = self._get_new_name(doc, newname)
            doc.getSheets().getByName(n).setName(new_name)
        return

    def sheetVisible(self, address, visible):
        doc = self._get_doc(address.Doc)
        names = self._get_sheets_names(address.Sheet, doc, False)
        for n in names:
            doc.getSheets().getByName(n).IsVisible = visible
        return

    def sheetPassword(self, address, password, remove):
        doc = self._get_doc(address.Doc)
        if address.Sheet:
            names = self._get_sheets_names(address.Sheet, doc, False)
        else:
            names = self.getSheetsNames(doc)
        for n in names:
            if remove:
                doc.getSheets().getByName(n).unprotect(password)
            else:
                doc.getSheets().getByName(n).protect(password)
        return

    def sheetSort(self, doc, asc):
        names = sorted(self.getSheetsNames(doc), reverse=not asc)
        for i, v in enumerate(names):
            doc.getSheets().moveByName(v, i)
        return

    def sheetCopy(self, address, newname, pos, rename):
        doc = self._get_doc(address.Doc)
        index = pos
        if pos < 0:
            index = doc.getSheets().getCount() + pos + 1
        if address.Sheet:
            if isinstance(address.Sheet, tuple):
                rename = True
            names = self._get_sheets_names(address.Sheet, doc, False)
        else:
            rename = True
            names = self.getSheetsNames(doc)
        for n in names:
            new_name = newname
            if rename:
                new_name = self._get_new_name(doc, newname)
            doc.getSheets().copyByName(n, new_name, index)
        return

    def sheetCopyToDoc(self, source, target, pos, rename, values):
        src_doc = self._get_doc(source.Doc)
        doc = target.Doc
        mode = 1
        if values:
            mode = 2
        index = pos
        if pos < 0:
            index = doc.getSheets().getCount() + pos + 1
        if source.Sheet:
            names = self._get_sheets_names(source.Sheet, src_doc, False)
        else:
            names = self.getSheetsNames(src_doc)
        for name in names:
            new_name = name
            if rename:
                new_name = self._get_new_name(doc, name)
            doc.getSheets().insertNewByName(new_name, index)
            sheet = doc.getSheets().getByName(new_name)
            sheet.link(src_doc.URL, name, '', '', mode)
            sheet.setLinkMode(0)
            doc.getCurrentController().setActiveSheet(sheet)
            dp = sheet.getDrawPage()
            src_dp = src_doc.getSheets().getByName(name).getDrawPage()
            src_doc.getCurrentController().setFormDesignMode(True)
            for i in range(src_dp.getCount()):
                src_img = src_dp.getByIndex(i)
                src_doc.getCurrentController().select(src_img)
                comun.copy(src_doc)
                new = comun.paste(doc)
                new.getByIndex(0).setPosition(src_img.getPosition())
            src_doc.getCurrentController().setFormDesignMode(False)
            doc.getCurrentController().setFormDesignMode(False)
        return

    def getCell(self, address):
        doc = self._get_doc(address.Doc)
        if address.Current:
            return doc.getCurrentSelection().getCellByPosition(0, 0)
        sheet = comun.get_sheet(doc, address.Sheet)
        if address.Name:
            cell = sheet.getCellRangeByName(address.Name.upper()).getCellByPosition(0, 0)
        else:
            cell = sheet.getCellByPosition(address.Col, address.Row)
        return cell

    def getRange(self, address):
        doc = self._get_doc(address.Doc)
        if address.Current:
            return doc.getCurrentSelection()
        sheet = comun.get_sheet(doc, address.Sheet)
        if address.Name:
            rango = sheet.getCellRangeByName(address.Name.upper())
        else:
            rango = sheet.getCellRangeByPosition(
                address.Col, address.Row, address.EndCol, address.EndRow)
        return rango

    def addRanges(self, container, ranges):
        if container.ImplementationName != 'ScCellRangesObj':
            container = container.createInstance('com.sun.star.sheet.SheetCellRanges')
        if not isinstance(ranges, tuple):
            ranges = (ranges,)
        for r in ranges:
            if not r:
                continue
            if not container.hasByName(r.AbsoluteName):
                container.addRangeAddress(r.getRangeAddress(), False)
        return container

    def removeRanges(self, container, ranges):
        if not isinstance(ranges, tuple):
            ranges = (ranges,)
        for r in ranges:
            if container.hasByName(r.AbsoluteName):
                container.removeRangeAddress(r.getRangeAddress())
        return container

    def getColumns(self, doc, ranges):
        container = doc.createInstance('com.sun.star.sheet.SheetCellRanges')
        if not isinstance(ranges, tuple):
            ranges = (ranges,)
        for r in ranges:
            for c in range(r.getColumns().getCount()):
                col = r.getColumns().getByIndex(c)
                container.addRangeAddress(col.getRangeAddress(), False)
        return container

    def getRows(self, doc, ranges):
        container = doc.createInstance('com.sun.star.sheet.SheetCellRanges')
        if not isinstance(ranges, tuple):
            ranges = (ranges,)
        for r in ranges:
            for r2 in range(r.getRows().getCount()):
                row = r.getRows().getByIndex(r2)
                container.addRangeAddress(row.getRangeAddress(), False)
        return container

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
        if not isinstance(data[0], tuple):
            data = tuple(zip(data))
        rango = comun.offset(cell, len(data[0]), len(data), True)
        rango.setDataArray(data)
        return

    def getData(self, cell, visible=False):
        rango = cell
        if comun.is_cell(cell):
            rango = self.getCurrentRegion(cell, False)
        if visible:
            data = ()
            ranges = self.getVisible(rango)
            for i in range(ranges.getCount()):
                data += ranges.getByIndex(i).getDataArray()
        else:
            data = rango.getDataArray()
        return data

    def getEmpty(self, cell):
        rango = cell
        if comun.is_cell(cell):
            rango = self.getCurrentRegion(cell, False)
        return rango.queryEmptyCells()

    def getVisible(self, cell):
        rango = cell
        if comun.is_cell(cell):
            rango = self.getCurrentRegion(cell, False)
        return rango.queryVisibleCells()


