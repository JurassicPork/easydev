# coding: utf-8

import logging
import datetime
import os

import uno
from com.sun.star.beans import PropertyValue, NamedValue
from com.sun.star.util import Time, Date, DateTime

from easydev.setting import (
    NAME_EXT,
)


log = logging.getLogger(NAME_EXT)



def replace_ext(path, ext):
    path, _, name, _ = get_path_info(path)
    return '{}/{}.{}'.format(path, name, ext)


def replace_filename(path, filename):
    path, _, _, _ = get_path_info(path)
    return os.path.join(path, filename)


def get_path_info(path):
    path = path_to_os(path)
    path, filename = os.path.split(path)
    name, extension = os.path.splitext(filename)
    return (path, filename, name, extension)


def basename(path):
    return os.path.basename(path)


def exists(path):
    return os.path.exists(path)


def path_to_os(path):
    if path.startswith('file://'):
        path = uno.fileUrlToSystemPath(path)
    return path


def path_to_url(path):
    if not path.startswith('file://'):
        path = uno.systemPathToFileUrl(path)
    return path


def set_properties(data):
    properties = []
    if isinstance(data[0], tuple):
        for p in data:
            pv = PropertyValue()
            pv.Name = p[0]
            pv.Value = p[1]
            properties.append(pv)
    elif isinstance(data[0], NamedValue):
        for p in data:
            pv = PropertyValue()
            pv.Name = p.Name
            pv.Value = p.Value
            properties.append(pv)
    return tuple(properties)


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

