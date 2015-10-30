Data Base
=========

ODBC
----

.. code-block:: vbnet

    Sub ConexionODBC()

        util = createUnoService("org.universolibre.EasyDev")

        odbc = "ConSQL"
        user = "sa"
        passw = "letmein"

        con = util.conODBC(odbc, user, passw)

        util.msgbox(con)

    End Sub
