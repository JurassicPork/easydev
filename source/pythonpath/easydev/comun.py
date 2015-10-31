# coding: utf-8

import logging
import datetime
import os

import uno
from com.sun.star.beans import PropertyValue, NamedValue
from com.sun.star.util import Time, Date, DateTime

from easydev.setting import (
    DATA_TYPES,
    NAME_EXT,
)


log = logging.getLogger(NAME_EXT)


class LODefault(object):

    def __init__(self, ctx, sm, desktop, toolkit):
        self.ctx = ctx
        self.sm = sm
        self.desktop = desktop
        self.toolkit = toolkit

    def _create_instance(self, name, with_context=True):
        if with_context:
            instance = self.sm.createInstanceWithContext(name, self.ctx)
        else:
            instance = self.sm.createInstance(name)
        return instance


def isdir(path):
    return os.path.isdir(path_to_os(path))


def replace_ext(path, ext):
    path, _, name, _ = get_path_info(path)
    return '{}/{}.{}'.format(path, name, ext)


def replace_filename(path, filename):
    path, _, _, _ = get_path_info(path)
    return os.path.join(path, filename)


def replace_name_ext(path_target, path_source, ext):
    _, _, name, _ = get_path_info(path_source)
    return os.path.join(path_target, '{}.{}'.format(path, name, ext))


def get_path_info(path):
    path = path_to_os(path)
    path, filename = os.path.split(path)
    name, extension = os.path.splitext(filename)
    return (path, filename, name, extension)


def basename(path):
    return os.path.basename(path_to_os(path))


def exists(path):
    return os.path.exists(path_to_os(path))


def path_to_os(path):
    if path.startswith('file://'):
        path = uno.fileUrlToSystemPath(path)
    return path


def path_to_url(path):
    if not path.startswith('file://'):
        path = uno.systemPathToFileUrl(path)
    return path


def set_properties(data):
    if not data:
        return ()
    properties = []
    if isinstance(data[0], tuple):
        for p in data:
            pv = PropertyValue()
            pv.Name = p[0]
            pv.Value = _verify_type(p[0], p[1])
            properties.append(pv)
    elif isinstance(data[0], NamedValue):
        for p in data:
            pv = PropertyValue()
            pv.Name = p.Name
            pv.Value = _verify_type(p.Name, p.Value)
            properties.append(pv)
    else:
        properties = data
    return tuple(properties)


def _verify_type(name, value):
    if name == 'FilterData':
        return uno.Any("[]com.sun.star.beans.PropertyValue", value)
    return value


def to_date(value):
    if isinstance(value, Time):
        new_value = datetime.time(value.Hours, value.Minutes, value.Seconds)
    elif isinstance(value, Date):
        new_value = datetime.date(value.Year, value.Month, value.Day)
    elif isinstance(value, DateTime):
        new_value = datetime.datetime(
            value.Year, value.Month, value.Day,
            value.Hours, value.Minutes, value.Seconds)
    else:
        new_value = value
    return new_value


def to_dict(data, test_date=False):
    if isinstance(data[0], tuple):
        if test_date:
            dic = {r[0]: to_date(r[1]) for r in data}
        else:
            dic = {r[0]: r[1] for r in data}
    elif isinstance(data[0], (NamedValue, PropertyValue)):
        if test_date:
            dic = {r.Name: to_date(r.Value) for r in data}
        else:
            dic = {r.Name: r.Value for r in data}
    return dic


def parse_data_type(resulset):
    if not resulset:
        return ()
    info = resulset.getMetaData()
    cols = range(1, info.getColumnCount() + 1)

    #~ for c in cols:
        #~ dt = info.getColumnTypeName(c).lower()
        #~ print ('DATA TYPE', dt)
        #~ if not DATA_TYPES.get(dt, False):
            #~ log.info(dt)

    cols_type = ('',) + tuple(DATA_TYPES[info.getColumnTypeName(c).lower()] for c in cols)
    headers = tuple(info.getColumnName(c) for c in cols)
    data = [headers]
    while resulset.next():
        row = tuple(getattr(resulset, cols_type[c])(c) for c in cols)
        data.append(row)
    return tuple(data)


