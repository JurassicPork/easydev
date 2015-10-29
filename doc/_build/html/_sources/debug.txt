Tools for debug
===============


Show info en shell
------------------

Only Linux

.. code-block:: vbnet

    Sub Debug()
        'Create service
        util = createUnoService("org.universolibre.EasyDev")

        util.debug("Test debug")
    End Sub

.. image:: images/img002.png
    :width: 500px
    :align: center

In Windows, you can open a new document Writer, save with name **debug.odt**
and show info debug in this document.

.. image:: images/img003.png
    :width: 500px
    :align: center


Save info in file
-----------------

.. code-block:: vbnet

    Sub LogFile()

        util = createUnoService("org.universolibre.EasyDev")

        util.log("/home/USER/log.txt", util)

    End Sub

Automatic add date and time. ::

    2015-10-28 20:56:35 - EasyDev - <uno_component.EasyDev object at 0x7f96caf34438>
    2015-10-28 20:56:41 - EasyDev - <uno_component.EasyDev object at 0x7f96caf34438>


Special msgbox
--------------

Show any data in msgbox like objects and arrays.

.. code-block:: vbnet

    Sub MessageBox()
        util = createUnoService("org.universolibre.EasyDev")

        'Show info in message box
        util.msgbox("Debug data")

        'Show any data
        data = "This is string"
        util.msgbox(data)

        data = 12345
        util.msgbox(data)

        data = Array("Uno", 2)
        util.msgbox(data)

        util.msgbox(util)
    End Sub


Call MRI
--------

MRI is the best extension for introspeccion of objects for Apache OpenOffice
and LibreOffice. `Download`_ and install.

Call from EasyDev.

.. code-block:: vbnet

    Sub CallMRI()
        util = createUnoService("org.universolibre.EasyDev")

        'MRI is a great extension
        util.mri(util)
    End Sub

.. image:: images/img004.png
    :width: 500px
    :align: center


.. _Download: http://extensions.openoffice.org/en/project/MRI
