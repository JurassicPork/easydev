# coding: utf-8

import logging
import datetime
import os
import tempfile
import zipfile
from functools import wraps

import uno
from com.sun.star.beans import PropertyValue, NamedValue
from com.sun.star.util import Time, Date, DateTime
from com.sun.star.awt import Size, Point

from easydev.setting import (
    DATA_TYPES,
    NAME_EXT,
    OBJECTS,
    PY2
)


if PY2:
    str = unicode

log = logging.getLogger(NAME_EXT)
CTX = uno.getComponentContext()
SM = CTX.getServiceManager()


def catch_exception(f):
    @wraps(f)
    def func(*args, **kwargs):
        try:
            return f(*args, **kwargs)
        except Exception as e:
            log.error(f.__name__, exc_info=True)
    return func


def _create_instance(name, with_context=True):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance


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
    return os.path.join(path_target, '{}.{}'.format(name, ext))


def get_path_info(path):
    path = path_to_os(path)
    path, filename = os.path.split(path)
    name, extension = os.path.splitext(filename)
    return (path, filename, name, extension)


def get_path_temp(name='', url=False):
    if name:
        temp = join(tempfile.gettempdir(), name)
    else:
        temp = tempfile.mkstemp()[1]
    if url:
        temp = path_to_url(temp)
    return temp


def kill(path):
    try:
        os.remove(path)
    except:
        pass
    return


def join(*paths):
    return os.path.join(*paths)


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
    if not data:
        return {}
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


def offset(cell, cols, rows, expand=False):
    sheet = cell.getSpreadsheet()
    cursor = sheet.createCursorByRange(cell)
    cursor.collapseToCurrentRegion()
    if expand:
        cursor.collapseToSize(
            cursor.Columns.getCount() + cols - 1, cursor.Rows.getCount() + rows - 1)
    else:
        cursor.collapseToSize(cols, rows)
    return sheet.getCellRangeByName(cursor.AbsoluteName)


def is_cell(cell):
    return cell.getImplementationName() == OBJECTS['CELL']


def copy(doc):
    dispatch = _create_instance('com.sun.star.frame.DispatchHelper')
    frame = doc.getCurrentController().getFrame()
    dispatch.executeDispatch(frame, '.uno:Copy', '', 0, ())
    return


def paste(doc):
    sc = _create_instance('com.sun.star.datatransfer.clipboard.SystemClipboard')
    transferable = sc.getContents()
    doc.getCurrentController().insertTransferable(transferable)
    return doc.getCurrentSelection()


def mri(obj):
    m = _create_instance('mytools.Mri')
    if m is None:
        return None
    m.inspect(obj)
    return True

def get_sheet(doc, data):
    if not data:
        return doc.getCurrentController().getActiveSheet()
    if isinstance(data, str):
        return doc.getSheets().getByName(data)
    if isinstance(data, int):
        index = data
        if index < 0:
            index = doc.getSheets().getCount() + index
        return doc.getSheets().getByIndex(index)
    return data

def get_pos_size(data):
    default = 500
    size = Size()
    pos = Point()
    if not data:
        size.Width = default
        size.Height = default
        return pos, size
    if data.Width:
        size.Width = data.Width
    else:
        size.Width = default
    if data.Height:
        size.Height = data.Height
    else:
        size.Height = default
    pos.X = data.X
    pos.Y = data.Y
    return pos, size


def set_property(control, properties):
    uno.invoke(control, 'setPropertyValue', properties)
    return


def zip_file(source, target):
    if not target:
        target = replace_ext(source, 'zip')
    elif isdir(target):
        target = replace_name_ext(target, source, 'zip')
    z = zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED)
    if exists(source):
        _, name, _, _ = get_path_info(source)
        z.write(source, name)
    z.close()
    return


def zip_dir(source, target):
    if not target:
        target = path_to_url(source).split('/')
        target[-1] = '{}.zip'.format(target[-1])
        target = path_to_os('/'.join(target))
    elif isdir(target):
        path = path_to_url(source).split('/')
        target = join(target, '{}.zip'.format(path[-1]))
    z = zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED)
    root_len = len(os.path.abspath(source))
    for root, dirs, files in os.walk(source):
        relative = os.path.abspath(root)[root_len:]
        for f in files:
            fullpath = join(root, f)
            file_name = join(relative, f)
            z.write(fullpath, file_name)
    z.close()
    return


def unzip(source, target, name):
    z = zipfile.ZipFile(source)
    if name:
        z.extract(name, target)
    else:
        if not target:
            path, _, name, _ = get_path_info(source)
            target = join(path, name)
        z.extractall(target)
    z.close()
    return

    #~ new_shape = doc.createInstance(SRV_GOS)
    #~ dp.add(new_shape)
    #~ new_shape.Graphic = src_img.Graphic
    #~ new_shape.setPosition(src_img.getPosition())
    #~ new_shape.setSize(src_img.getSize())
