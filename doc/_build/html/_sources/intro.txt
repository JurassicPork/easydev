Introduction
============
EasyDev can you help to develop macros in LibreOffice.

Installation
------------
    * Download lates version from `GitHub`_.
    * Open LibreOffice
    * Tools -> Extension manager...
    * Add...

.. image:: images/img001.png
    :width: 500px
    :align: center

Test installation
-----------------

.. code-block:: vbnet

    Sub Test()
        'Create service
        util = createUnoService("org.universolibre.EasyDev")

        MsgBox util.VERSION
    End Sub






.. _GitHub: https://github.com/UniversoLibreMexicoAC/easydev/tree/master/files
