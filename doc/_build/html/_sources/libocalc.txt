Calc
===============

Get sheet
---------

Get all sheets names.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")

    'Get current doc
    doc = util.getDoc("")
    names = util.getSheetsNames(doc)
    util.msgbox(names)

Get active sheet.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")

    'Get current doc
    doc = util.getDoc("")
    sheet = util.getSheet(doc, "")
    util.msgbox(sheet.getName())

Get sheet by name.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")

    'Get current doc
    doc = util.getDoc("")
    sheet = util.getSheet(doc, "Sheet1")
    util.msgbox(sheet.getName())

Get sheet by index. Index start in 0.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")

    doc = util.getDoc("")
    sheet = util.getSheet(doc, 0)
    util.msgbox(sheet.getName())

Get last sheet by index.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")

    doc = util.getDoc("")
    sheet = util.getSheet(doc, -1)
    util.msgbox(sheet.getName())


Insert sheets
-------------

Insert sheet in pos.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")

    doc = ThisComponent
    new_name = "New Sheet"
    util.sheetInsert(doc, new_name, 1)

Insert in last pos.

.. code-block:: vbnet

    new_name = "New Sheet 2"
    util.sheetInsert(doc, new_name, -1)


Remove sheet
------------

Remove sheet by name

.. code-block:: vbnet

    sheet_name = "Sheet1"
    util.sheetRemove(doc, sheet_name)

Remove by object

.. code-block:: vbnet

    'Get active sheet
    obj_sheet = util.getSheet(doc, "")
    util.sheetRemove(doc, obj_sheet)


Move sheet
----------

Move sheet by name to pos.

.. code-block:: vbnet

    sheet_name = "Sheet3"
    util.sheetMove(doc, sheet_name, 0)

Move sheet by object to last pos.

.. code-block:: vbnet

    'Get active sheet
    sheet = util.getSheet(doc, "")
    util.sheetMove(doc, sheet, -1)


Sort sheets
-----------

Sort sheets by name.

.. code-block:: vbnet

    util.sheetSort(ThisComponent, True)

Sort in reverse.

.. code-block:: vbnet

    util.sheetSort(ThisComponent, False)


Copy sheet
----------

Copy sheet by name. Get error if new name exists.

.. code-block:: vbnet

    doc = ThisComponent
    sheet_name = "Sheet1"
    new_name = "New Sheet"
    util.sheetCopy(doc, sheet_name, new_name, 0, False)

Copy sheet by name, force rename if new sheet name exists.

.. code-block:: vbnet

    doc = ThisComponent
    sheet_name = "Sheet1"
    new_name = "New Sheet"
    util.sheetCopy(doc, sheet_name, new_name, 0, True)

Copy sheet by object in last pos and force rename.

.. code-block:: vbnet

    'Get active sheet
    sheet = util.getSheet(doc, "")
    new_name = "New Sheet"
    util.sheetCopy(doc, sheet, new_name, -1, True)


Cells
-----

Get active cell, always back only one cell

.. code-block:: vbnet

    Sub getCellDoc()
        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        address.Current = True
        cell = util.getCell(address)
        msg = util.format("{} {}", Array(cell.ImplementationName, cell.AbsoluteName))
        util.msgbox(msg)
    End Sub


Get active cell from document by title name, document must be open

.. code-block:: vbnet

        address.Doc = "test.ods"    'Title doc
        address.Current = True
        cell = util.getCell(address)
        msg = util.format("{} {}", Array(cell.ImplementationName, cell.AbsoluteName))
        util.msgbox(msg)

Get active cell from object doc

.. code-block:: vbnet

        doc = util.getDoc("")
        address.Doc = doc
        address.Current = True
        cell = util.getCell(address)
        msg = util.format("{} {}", Array(cell.ImplementationName, cell.AbsoluteName))
        util.msgbox(msg)

Get cell from object doc, sheet by name, cell by name

.. code-block:: vbnet

    doc = util.getDoc("")
    address.Doc = doc
    address.Sheet = "Sheet2"
    address.Name = "B5"
    cell = util.getCell(address)
    msg = util.format("{} {}", Array(cell.ImplementationName, cell.AbsoluteName))
    util.msgbox(msg)

Get cell from object doc, sheet by name, cell by col, row

.. code-block:: vbnet

    doc = util.getDoc("")
    address.Doc = doc
    address.Sheet = "Sheet2"
    address.Name = ""
    address.Col = 1
    address.Row = 4
    cell = util.getCell(address)
    msg = util.format("{} {}", Array(cell.ImplementationName, cell.AbsoluteName))
    util.msgbox(msg)

Get cell from object sheet, object doc is ommited

.. code-block:: vbnet

    sheet = doc.getCurrentController().getActiveSheet()
    address.Sheet = sheet
    address.Name = "D5"
    cell = util.getCell(address)
    msg = util.format("{} {}", Array(cell.ImplementationName, cell.AbsoluteName))
    util.msgbox(msg)

Ranges
------

Get active range

.. code-block:: vbnet

    Sub getRangeDoc()
        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        address.Current = True
        range = util.getRange(address)
        msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
        util.msgbox(msg)
    End Sub

Get active range from document by title name, document must be open

.. code-block:: vbnet

    address.Doc = "test.ods"    'Title doc
    address.Current = True
    range = util.getRange(address)
    msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
    util.msgbox(msg)

Get active range from object doc

.. code-block:: vbnet

        doc = util.getDoc("")
        address.Doc = doc
        address.Current = True
        range = util.getRange(address)
        msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
        util.msgbox(msg)

Get range from object doc, sheet by name, range by name

.. code-block:: vbnet

    doc = util.getDoc("")
    address.Doc = doc
    address.Sheet = "Sheet2"
    address.Name = "B5:C10"
    range = util.getRange(address)
    msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
    util.msgbox(msg)

Get range from object doc, sheet by name, range by position

.. code-block:: vbnet

    address.Sheet = "Sheet2"
    address.Name = ""
    address.Col = 1
    address.Row = 4
    address.EndCol = 3
    address.EndRow = 9
    range = util.getRange(address)
    msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
    util.msgbox(msg)

Get range from object sheet, object doc is ommited

.. code-block:: vbnet

    sheet = doc.getCurrentController().getActiveSheet()
    address.Sheet = sheet
    address.Name = "D5:E10"
    range = util.getRange(address)
    msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
    util.msgbox(msg)

Select range
------------

.. code-block:: vbnet

    Sub SelectRange()
        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        doc = util.getDoc("")
        address.Doc = doc
        address.Sheet = "Sheet2"
        address.Name = "B5:C10"
        range = util.getRange(address)

        'Select
        util.selectRange(doc, range)

    End Sub

Current region
--------------

.. code-block:: vbnet

    Sub getCurrentRegion()

        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        address.Current = True
        cell = util.getCell(address)
        msg = util.format("{} {}", Array(cell.ImplementationName, cell.AbsoluteName))
        util.msgbox(msg)

    End Sub

Only one cell into current region is necesary, get cursor

.. code-block:: vbnet

    cursor = util.getCurrentRegion(cell, True)
    msg = util.format("{} {}", Array(cursor.ImplementationName, cursor.AbsoluteName))
    util.msgbox(msg)

Only one cell into current region is necesary, get range

.. code-block:: vbnet

    range = util.getCurrentRegion(cell, False)
    msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
    util.msgbox(msg)


Get empty cells
---------------

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    'Get active cell, always get one cell
    address.Current = True
    cell = util.getCell(address)

    ranges = util.getEmpty(cell)
    util.selectRange(ThisComponent, ranges)

Get visible cells
-----------------

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    'Get active cell, always get one cell
    address.Current = True
    cell = util.getCell(address)

    ranges = util.getVisible(cell)
    util.selectRange(ThisComponent, ranges)


Last row
--------

Get last row into current region.

.. code-block:: vbnet

    Sub getLastRow()
        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        address.Current = True
        cell = util.getCell(address)
        row = util.getLastRow(cell)
        util.msgbox(row)
    End Sub

Next value
----------

In range values. Get nex value in relative column (0), or (1), etc.

.. code-block:: vbnet

    Sub getNextID()
        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        address.Current = True
        cell = util.getCell(address)
        value = util.getNextID(cell, 0)
        util.msgbox(value)
    End Sub

.. image:: images/img009.png
    :width: 400px
    :align: center


Cell value
----------

Automatically detect type: string, value or formula.

.. code-block:: vbnet

    Sub CellValue()
        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        address.Current = True
        cell = util.getCell(address)

        value = "String"
        util.setValue(cell, value)
        value = util.getValue(cell)
        util.msgbox(value)

        value = 12345
        util.setValue(cell, value)
        value = util.getValue(cell)
        util.msgbox(value)

        value = "=SUM(A1:C1)"
        util.setValue(cell, value)
        value = util.getValue(cell)
        util.msgbox(value)

    End Sub


.. _setdata:

Set data
--------

Automatically calculate width and height of data size.

.. code-block:: vbnet

    Sub SetData()
        util = createUnoService("org.universolibre.EasyDev")
        address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

        'Get active cell, always get one cell
        address.Current = True
        cell = util.getCell(address)

        data = Array( _
            Array(1, "Uno", "Tres"), _
            Array(2, "Dos", "Cuatro"), _
            Array(3, "Tres", "Cinco"), _
        )
        'Automatically calculate width and height of data size
        util.setData(cell, data)

    End Sub


Get data
--------

Get data array from cell, automatically get current region. Get hidden cells inclusive.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    'Get active cell, always get one cell
    address.Current = True
    cell = util.getCell(address)

    data = util.getData(cell, False)
    util.msgbox(data)

Get data array, only visible cells.

.. code-block:: vbnet

    data = util.getData(cell, True)
    util.msgbox(data)

