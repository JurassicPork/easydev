Writer
======

Paragraphs
----------

Get all paragraphs in document.

.. code-block:: vbnet

    Sub GetParagraps()
        util = createUnoService("org.universolibre.EasyDev")

        'Get current doc
        doc = util.getDoc("")

        'Get all paragraphs
        paragraphs = util.getParagraphs(doc, True)
        util.msgbox(util.len(paragraphs))

    End Sub

Get paragraphs with text

.. code-block:: vbnet

    paragraphs = util.getParagraphs(doc, False)
    util.msgbox(util.len(paragraphs))


