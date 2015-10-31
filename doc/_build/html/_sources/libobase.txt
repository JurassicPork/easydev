Data Base
=========

ODBC
----

I test connection with: MySQL, PostgreSQL, SQLite and MSSQL.

.. code-block:: vbnet

    Sub ConexionODBC()

        util = createUnoService("org.universolibre.EasyDev")

        odbc = "ConSQL"
        user = "sa"
        passw = "letmein"

        con = util.conODBC(odbc, user, passw)

        util.msgbox(con)

    End Sub

.. _base-query:

Query
-----

Make query, get data like array

.. code-block:: vbnet

    odbc = "ODBCSQLITE"
    user = ""
    passw = ""
    con = util.conODBC(odbc, user, passw)

    sql = "SELECT * FROM contactos"
    data = util.query(con, sql, True)
    util.msgbox(data)

Make query, get data like resulset

.. code-block:: vbnet

    sql = "SELECT * FROM contactos"
    data = util.query(con, sql, False)
    util.msgbox(data)

You can set resulset to grid, see :ref:`grid`.
