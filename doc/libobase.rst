Data Base
=========

Connection DB
-------------

Connect to register DB in Base.

.. code-block:: vbnet

    Sub ConnectDB()
        db_name = "test"
        user = ""
        pass = ""
        con = util.conDB(db_name, user, pass)

        util.msgbox(con.isClosed())
    End Sub

Verify if exists.

.. code-block:: vbnet

    db_name = "test"
    MsgBox util.existsDB(db_name)

If exists, get path location.

.. code-block:: vbnet

    db_name = "test"
    If util.existsDB(db_name) Then
        MsgBox util.getPathDB(db_name)
    End If

Create DB and register in Base.

.. code-block:: vbnet

    path_db = "/home/USER/dbtest.odb"
    db_name = "TestOne"
    util.newDB(path_db)
    util.registerDB(db_name, path_db)

Revoke register DB.

.. code-block:: vbnet

    db_name = "test"
    util.revokeDB(db_name)


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
