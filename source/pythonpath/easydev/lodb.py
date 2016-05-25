# coding: utf-8

import logging

import apsw

from org.universolibre.EasyDev import XLODataBase
from easydev import comun
from easydev.comun import LODefault
from easydev.setting import LOG, NAME_EXT, PY2

if PY2:
    str = unicode

log = logging.getLogger(NAME_EXT)


class LODataBase(XLODataBase, LODefault):


    def __init__(self, ctx, sm, desktop, toolkit):
        LODefault.__init__(self, ctx, sm, desktop, toolkit)
        self._dbc = self._create_instance('com.sun.star.sdb.DatabaseContext')

    def conODBC(self, name, user, password):
        dm = self._create_instance('com.sun.star.sdbc.DriverManager')
        pv = comun.set_properties((('user', user), ('password', password)))
        try:
            con = dm.getConnectionWithInfo('sdbc:odbc:{}'.format(name), pv)
            return con
        except Exception as e:
            log.error(str(e))
            return None

    def registerDB(self, name, path):
        self._dbc.registerDatabaseLocation(name, comun.path_to_url(path))
        return True

    def revokeDB(self, name):
        self._dbc.revokeDatabaseLocation(name)
        return True

    def existsDB(self, name):
        return self._dbc.hasRegisteredDatabase(name)

    def getPathDB(self, name):
        return comun.path_to_os(self._dbc.getDatabaseLocation(name))

    def conDB(self, name, user, password):
        db = self._dbc.getByName(name)
        con = db.getConnection(user, password)
        return con

    def query(self, con, sql, array):
        cursor = con.createStatement()
        resulset = cursor.executeQuery(sql)
        if array:
            return comun.parse_data_type(resulset)
        else:
            return resulset

    def update(self, con, sql):
        cursor = con.createStatement()
        rows = cursor.executeUpdate(sql)
        return rows

    def sqliteExecute(self, pathdb, sql):
        try:
            connection = apsw.Connection(pathdb)
            with connection as db:
                db.cursor().execute(sql)
            return True
        except Exception as e:
            log.error(str(e))
            return False

    def sqliteSelect(self, pathdb, sql):
        connection = apsw.Connection(pathdb)
        with connection as db:
            result = db.cursor().execute(sql)
        return tuple(result)


