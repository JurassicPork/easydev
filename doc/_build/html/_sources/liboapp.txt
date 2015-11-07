Documents
=========

New
---

Look: `<http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XComponentLoader.html>`_

.. code-block:: vbnet

    Sub NewDoc()
        'Create new doc
        util = createUnoService("org.universolibre.EasyDev")

        'Default Calc
        doc = util.newDoc("")

    End Sub

Other values: swriter, simpress, sdraw, smath

.. code-block:: vbnet

        doc = util.newDoc("sdraw")

For Base document.

.. code-block:: vbnet

    path_db = "/home/USER/dbtest.odb"
    db = util.newDB(path_db)


Get doc
-------

Current

.. IMPORTANT::
   Current doc can be IDE

.. code-block:: vbnet

    doc1 = util.getDoc("")
    MsgBox doc1.Title

Get doc by title, if not found, doc is Empty

.. code-block:: vbnet

    doc2 = util.getDoc("Name_doc.odt")
    util.msgbox(doc2)


Get type
--------

Values return: calc, writer, impress, draw, math, base, ide

.. code-block:: vbnet

    doc1 = util.newDoc("sdraw")
    MsgBox util.getTypeDoc(doc1)

    doc2 = util.newDoc("swriter")
    MsgBox util.getTypeDoc(doc2)


Get documents
-------------

Get all open documents

.. code-block:: vbnet

    docs = util.getDocs()
    MsgBox util.format("{} Open documents", util.len(docs))


Open
----

More info and options:

    * `Component Loader <http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html>`_
    * `Media Descriptor <http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1MediaDescriptor.html>`_

Open by path document

.. code-block:: vbnet

    Dim options1(0) As New com.sun.star.beans.NamedValue
    util = createUnoService("org.universolibre.EasyDev")

    path = "/home/USER/Plantilla.ods"
    options = Array()
    doc = util.openDoc(path, options)

Open like template

.. code-block:: vbnet

    options1(0).Name = "AsTemplate"
    options1(0).Value = True
    path = "/home/USER/Plantilla.ods"
    doc = util.openDoc(path, options1)

Open hidden

.. code-block:: vbnet

    options1(0).Name = "Hidden"
    options1(0).Value = True
    path = "/home/USER/Plantilla.ods"
    doc = util.openDoc(path, options1)
    MsgBox "Close doc"
    doc.dispose()


Active
------

Send focus to document

.. code-block:: vbnet

    doc1 = util.newDoc("")
    doc2 = util.newDoc("swriter")
    wait(1000)
    util.setFocus(doc1)


Status bar
----------

Set text and show progress bar

.. code-block:: vbnet

    'Get current doc
    doc = util.getDoc("")
    'Get status bar
    sb = util.getStatusBar(doc)

    'Init text and up limit
    sb.start( "Row ", 10 )
    For co1 = 1 To 10
        'Set value
        sb.setValue( co1 )
        Wait 1000
    Next
    'Is import free status bar
    sb.end()


Set visible
-----------

Hidden document.

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")

    doc = util.newDoc("")

    util.setVisible(doc, False)

    MsgBox "Document is hidden"

    util.setVisible(doc, True)


Export PDF
----------

All options in `PDF Export <http://wiki.services.openoffice.org/wiki/API/Tutorials/PDF_export>`_ in wiki.

If export is correct, return path save PDF

Export current doc in the same folder and same name

.. code-block:: vbnet

    doc = util.getDoc("")
    path = util.exportPDF(doc, "", Array())
    MsgBox util.format("PDF export in: {}", path)

For save in other folder and same name

.. code-block:: vbnet

    path_save = "/home/USER/OTHER_FOLDER"
    path_pdf = util.exportPDF(doc, path_save, Array())

For save in other folder and other name

.. code-block:: vbnet

    path_save = "/home/USER/OTHER_FOLDER/NAME.pdf"
    path_pdf = util.exportPDF(doc, path_save, Array())

Export with options

.. code-block:: vbnet

    Dim options(0) As New com.sun.star.beans.NamedValue

    doc = util.getDoc("")
    options(0).Name = "PageRange"
    options(0).Value = "2"
    path = util.exportPDF(doc, "", options)
    MsgBox util.format("PDF export in: {}", path)
