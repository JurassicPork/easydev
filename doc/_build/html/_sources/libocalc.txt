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

.. IMPORTANT::
   Current doc can be IDE, and IDE not have sheets

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    'Get active sheet from active doc
    sheet = util.getSheet(address)
    msgbox sheet.Name

Is better use explicit doc

.. code-block:: vbnet

    doc = ThisComponent
    address.Doc = doc
    sheet = util.getSheet(address)
    MsgBox sheet.Name

Get active sheet from other doc

.. code-block:: vbnet

    address.Doc = util.getDoc("Untitled 3.ods")
    sheet = util.getSheet(address)
    MsgBox sheet.Name

Get sheet by name.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = "Sheet2"
    sheet = util.getSheet(address)
    MsgBox sheet.Name

Get sheet by index. Index start in 0.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 1
    sheet = util.getSheet(address)
    MsgBox sheet.Name

Get last sheet by index.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = -1
    sheet = util.getSheet(address)
    MsgBox sheet.Name


Set active sheet
----------------

Active sheet by name.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    doc = ThisComponent
    address.Doc = doc
    address.Sheet = "Sheet2"

    util.sheetActivate(address)

Active sheet by index. Last sheet index == -1

.. code-block:: vbnet

    doc = ThisComponent
    address.Doc = doc
    address.Sheet = -1

    util.sheetActivate(address)


Insert sheets
-------------

Insert sheet in pos. Get error if sheet exists

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "New Sheet"
    util.sheetInsert(address, 0, False)

Rename sheet if exists. Get name "New Sheet_1", "New Sheet_2", etc...

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = "New Sheet"
    util.sheetInsert(address, 0, True)

Insert in last pos.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = "New Sheet"
    util.sheetInsert(address, -1, True)

Insert more one sheet.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = Array("One", "Two", "Other")
    util.sheetInsert(address, 0, True)


Insert sheet from document
--------------------------

Only LibreOffice, for OpenOffice use: :ref:`copysheet`.
Insert all sheets from source document to target document.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    source = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")
    target = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    doc = ThisComponent
    source.Doc = doc

    new_doc = util.newDoc("")
    target.Doc = new_doc

    util.sheetInsertFromDoc(source, target, -1, False)


Remove sheet
------------

.. IMPORTANT::
   The index of sheets change if remove some sheet.

Remove sheet by name

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "Sheet2"
    util.sheetRemove(address, False)

Remove sheet by index

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 1
    util.sheetRemove(address, False)

Remove sheet by index, last sheet

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = -1
    util.sheetRemove(address, False)

First get sheet and remove by object

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = "Sheet6"
    sheet = util.getSheet(address)
    address.Sheet = sheet
    util.sheetRemove(address, False)

Remove more one sheet, by index, name or object

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = Array(0, "Sheet4", "Sheet2", sheet, -1)
    util.sheetRemove(address, False)

Remove all sheet, except name

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = "Sheet3"
    util.sheetRemove(address, True)

Remove all sheet, except more one

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = Array("Sheet2", "Sheet6")
    util.sheetRemove(address, True)


Move sheet
----------

Move sheet by name to pos.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "Sheet2"
    util.sheetMove(address, 0)

Move first sheet to last pos

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 0
    util.sheetMove(address, -1)

Move sheet by object

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 1
    sheet = util.getSheet(address)
    address.Sheet = sheet
    util.sheetMove(address, -1)

Move more one sheet

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = Array ("Sheet2", "Sheet4")
    util.sheetMove(address, -1)


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

.. IMPORTANT::
   if you not use rename argument, you first valid if new name sheet not exists
   with method: hasByName

Copy sheet by name to last pos.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "Sheet2"
    util.sheetCopy(address, "NewName", -1, False)

Copy sheet by index to pos.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 0
    util.sheetCopy(address, "OtherName", -1, False)

Copy sheet by index to pos and rename if exists

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 0
    util.sheetCopy(address, "NameX", -1, True)

Copy more one sheet, rename is ommited, always is true

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = Array("Sheet1", "Sheet2")
    util.sheetCopy(address, "NameY", -1, True)

Copy all sheets, rename is ommited, always is true

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = ""
    util.sheetCopy(address, "NameNext", -1, False)


.. _copysheet:

Copy sheet to document
----------------------

.. IMPORTANT::
   The origin document must be saved before copy sheet

::

    sheetCopyToDoc(SOURCE, TARGET, POSITION, RENAME, BY_VALUE)

Source and targe are struct ``org.universolibre.EasyDev.CellRangeAddress``

* **SOURCE**: Document and sheet origin
* **TARGET**: Target document
* **POSITION**: Target position for example first position (0) or last position (-1)
* **RENAME**: If exists sheet name in target, rename sheet origin
* **BY_VALUE**: False, copy by formula, True, copy by values.

Copy all sheets from source to target document.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    source = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")
    target = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    doc = ThisComponent
    source.Doc = doc

    new_doc = util.newDoc("")
    target.Doc = new_doc

    util.sheetCopyToDoc(source, target, -1, True, False)

Copy sheet by name.

.. code-block:: vbnet

    doc = ThisComponent
    source.Doc = doc
    source.Sheet = "Source"
    new_doc = util.newDoc("")
    target.Doc = new_doc

    util.sheetCopyToDoc(source, target, -1, True, False)

Copy sheet by index

.. code-block:: vbnet

    doc = ThisComponent
    source.Doc = doc
    source.Sheet = 1
    new_doc = util.newDoc("")
    target.Doc = new_doc

    util.sheetCopyToDoc(source, target, -1, True, False)

Copy more one sheet

.. code-block:: vbnet

    doc = ThisComponent
    source.Doc = doc
    source.Sheet = Array(0, "Source")
    new_doc = util.newDoc("")
    target.Doc = new_doc

    util.sheetCopyToDoc(source, target, -1, True, False)

Copy sheet by value

.. code-block:: vbnet

    doc = ThisComponent
    source.Doc = doc
    source.Sheet = "Source"
    new_doc = util.newDoc("")
    target.Doc = new_doc

    util.sheetCopyToDoc(source, target, -1, True, True)


Rename sheet
------------

Rename sheet by name.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "Sheet2"
    util.sheetName(address, "Two")

Rename sheet by index.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 0
    util.sheetName(address, "First")

Rename more one sheet.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = Array(0, "Sheet3")
    util.sheetName(address, "NewName")


Show or hidden sheet
--------------------

.. IMPORTANT::
   If you try hide all sheets, not get error, but always at least one sheet will be visible

Hidden sheet by name.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "Sheet2"
    util.sheetVisible(address, False)

Hidden first sheet.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = 0
    util.sheetVisible(address, False)

Hidden last sheet.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = -1
    util.sheetVisible(address, False)

Hidden all sheets, except one

.. code-block:: vbnet

    sheets = util.getSheetsNames(ThisComponent)
    'First show sheet
    address.Doc = ThisComponent
    address.Sheet = sheets(0)
    util.sheetVisible(address, True)
    'Remove sheet from array
    sheets = util.delete(sheets, 0)
    'And hide the rest of sheets
    address.Sheet = sheets
    util.sheetVisible(address, False)

Show al sheets

.. code-block:: vbnet

    sheets = util.getSheetsNames(ThisComponent)
    address.Doc = ThisComponent
    address.Sheet = sheets
    util.sheetVisible(address, True)


Sheet password
--------------

Set password by name.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "Sheet2"
    util.sheetPassword(address, "letmein", False)

Set password by index.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = -1
    util.sheetPassword(address, "letmein", False)

Set password all sheet.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = ""
    util.sheetPassword(address, "letmein", False)

Remove password all sheet.

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = ""
    util.sheetPassword(address, "letmein", True)


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


.. _getranges:

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

Get range from object sheet, object doc is ommited.

.. code-block:: vbnet

    sheet = doc.getCurrentController().getActiveSheet()
    address.Sheet = sheet
    address.Name = "D5:E10"
    range = util.getRange(address)
    msg = util.format("{} {}", Array(range.ImplementationName, range.AbsoluteName))
    util.msgbox(msg)

Add ranges into container.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    address = createUnoStruct("org.universolibre.EasyDev.CellRangeAddress")

    address.Doc = ThisComponent
    address.Sheet = "Sheet1"
    address.Name = "A1:B2"
    range = util.getRange(address)
    'Create container, first argument is a document
    container = util.addRanges(ThisComponent, range)
    MsgBox container.getRangeAddressesAsString()

    'Add range to container
    address.Name = "A4:B5"
    range = util.getRange(address)
    'First argument is a container, previously created
    container = util.addRanges(container, range)
    MsgBox container.getRangeAddressesAsString()

Initial container from more one cell range

.. code-block:: vbnet

    address.Doc = ThisComponent
    address.Sheet = "Sheet1"
    address.Name = "B1:C2"
    r1 = util.getRange(address)
    address.Name = "B4:C4"
    r2 = util.getRange(address)
    address.Name = "D6:E6"
    r3 = util.getRange(address)
    ranges = Array(r1, r2, r3)

    'Create container, first argument is a document
    container = util.addRanges(ThisComponent, ranges)
    MsgBox container.getRangeAddressesAsString()

Remove ranges from container.

.. code-block:: vbnet

    r = util.removeRanges(container, r2)
    MsgBox container.getRangeAddressesAsString()

Get all columns from cell range, get a container.

.. code-block:: vbnet

    doc = ThisComponent

    address.Doc = doc
    address.Sheet = "Sheet1"
    address.Name = "B1:E1"
    range = util.getRange(address)

    columns = util.getColumns(ThisComponent, range)

    util.selectRange(doc, columns)

Get all rows from cell range, get a container.

.. code-block:: vbnet

    doc = ThisComponent

    address.Doc = doc
    address.Sheet = "Sheet1"
    address.Name = "A5:A10"
    range = util.getRange(address)

    rows = util.getRows(ThisComponent, range)

    util.selectRange(doc, rows)


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

