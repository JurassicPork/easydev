# coding: utf-8

import logging
from org.universolibre.EasyDev import XLODataBase
from easydev.setting import LOG, NAME_EXT
from easydev import comun
from easydev.comun import LODefault


log = logging.getLogger(NAME_EXT)


class LODataBase(XLODataBase, LODefault):
    types = {
        'integer': 'getLong',
        'varchar': 'getString',
    }

    def __init__(self, ctx, sm, desktop, toolkit):
        LODefault.__init__(self, ctx, sm, desktop, toolkit)

    def conODBC(self, name, user, password):
        dm = self._create_instance('com.sun.star.sdbc.DriverManager')
        pv = comun.set_properties((('user', user), ('password', password)))
        try:
            con = dm.getConnectionWithInfo('sdbc:odbc:{}'.format(name), pv)
            return con
        except Exception as e:
            log.error(str(e))
            return None

    def query(self, con, sql, array):
        cursor = con.createStatement()
        resulset = cursor.executeQuery(sql)
        if array:
            return self._parse_data(resulset)
        else:
            return resulset

    def _parse_data(self, r):
        if not r:
            return ()
        info = r.getMetaData()
        cols = range(1, info.getColumnCount() + 1)
        cols_type = ('',) + tuple(self.types[info.getColumnTypeName(c)] for c in cols)
        headers = tuple(info.getColumnName(c) for c in cols)
        data = [headers]
        while r.next():
            row = tuple(getattr(r, cols_type[c])(c) for c in cols)
            data.append(row)
        return tuple(data)

