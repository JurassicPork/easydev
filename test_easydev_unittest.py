#!/usr/bin/env python3
# coding: utf-8

import uno
import unittest
from unittest.mock import patch
import tempfile
import os
import subprocess
import datetime
import time
from io import StringIO

from com.sun.star.beans import NamedValue
from com.sun.star.beans import PropertyValue


class LIBO(object):

    ARG = 'socket,host=localhost,port=8100;urp;StarOffice.ComponentContext'
    LIBO = None
    CTX = None
    SM = None

    def __init__(self):
        self._check_process()
        local_context = uno.getComponentContext()
        local_resolver = local_context.ServiceManager.createInstanceWithContext(
            'com.sun.star.bridge.UnoUrlResolver', local_context )
        try:
            self.CTX = local_resolver.resolve('uno:{}'.format(self.ARG))
            self.SM = self.CTX.ServiceManager
            self.desktop = self.create_instance('com.sun.star.frame.Desktop', True)
        except Exception as e:
            pass

    def _check_process(self):
        proc = subprocess.Popen(['pgrep', 'soffice'], stdout=subprocess.PIPE)
        if not proc.communicate()[0]:
            self.LIBO = subprocess.Popen([
                'soffice', '--headless', '--accept={}'.format(self.ARG)])
            time.sleep(2)
        return

    def __del__(self):
        if self.LIBO:
            self.desktop.terminate()
            self.LIBO.terminate()
        time.sleep(1)

    def create_instance(self, name, with_context=False):
        if with_context:
            instance = self.SM.createInstanceWithContext(name, self.CTX)
        else:
            instance = self.SM.createInstance(name)
        return instance

    def _switch_path(self, path):
        if path.startswith('file://'):
            return uno.fileUrlToSystemPath(path)
        else:
            return uno.systemPathToFileUrl(path)

    def new_doc(self, type_doc='scalc'):
        path = 'private:factory/{}'.format(type_doc)
        doc = self.desktop.loadComponentFromURL(path, '_default', 0, ())
        return doc

    def open_doc(self, path):
        path_url = _switch_path(path)
        doc = self.desktop.loadComponentFromURL(path_url, '_blank', 0, ())
        return doc


class TestTools(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.libo = LIBO()
        self.tools = self.libo.create_instance('org.universolibre.EasyDev')

    def tearDown(self):
        del self.libo

    #~ def test_debug(self):
        #~ expected = 'Test debug'
        #~ with patch('sys.stdout', new=StringIO()) as fake_out:
            #~ self.tools.debug(expected) # Wrong
            #~ print (expected) # OK
            #~ result = fake_out.getvalue().strip()
            #~ self.assertEqual(result, expected)

    def test_get_size_screen(self):
        expected = '1366x768'
        result = self.tools.getSizeScreen()
        self.assertEqual(result, expected)

    def test_get_info_pc(self):
        expected = 'mau'
        result = self.tools.getInfoPC()[0]
        self.assertEqual(result, expected)

    def test_question(self):
        result = self.tools.question('Test', 'Test')
        self.assertFalse(result)

    def test_render_tuple(self):
        expected = 'Nombre: Teresa\nPaís: Portugal'
        template = 'Nombre: $name\nPaís: $country'
        data = (
            ('name', 'Teresa'),
            ('country', 'Portugal')
        )
        result = self.tools.render(template, data)
        self.assertEqual(result, expected)

    def test_render_named_value(self):
        expected = 'Nombre: Teresa\nPaís: Portugal'
        template = 'Nombre: $name\nPaís: $country'
        data = (
            NamedValue('name', 'Teresa'),
            NamedValue('country', 'Portugal'),
        )
        result = self.tools.render(template, data)
        self.assertEqual(result, expected)

    def test_render_property_value(self):
        expected = 'Nombre: Teresa\nPaís: Portugal'
        template = 'Nombre: $name\nPaís: $country'
        data = (
            PropertyValue('name', -1, 'Teresa', 1),
            PropertyValue('country', -1, 'Portugal', 1),
        )
        result = self.tools.render(template, data)
        self.assertEqual(result, expected)

    def test_render_number(self):
        expected = 'Age: 40\nPhone: 12345678'
        template = 'Age: $age\nPhone: $phone'
        data = (
            ('age', 40),
            ('phone', 12345678)
        )
        result = self.tools.render(template, data)
        self.assertEqual(result, expected)

    def test_format_string(self):
        expected = 'Number: one'
        template = 'Number: {}'
        data = 'one'
        result = self.tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_number(self):
        expected = '123'
        template = '{:d}'
        data = 123
        result = self.tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_list(self):
        expected = 'Numbers: one and two'
        template = 'Numbers: {} and {}'
        data = ('one', 'two')
        result = self.tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_list_pos(self):
        expected = 'Numbers: two and one'
        template = 'Numbers: {1} and {0}'
        data = ('one', 'two')
        result = self.tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_dict(self):
        expected = 'Numbers: one and two'
        template = 'Numbers: {n1} and {n2}'
        data = (('n1', 'one'), ('n2', 'two'))
        result = self.tools.format(template, data)
        self.assertEqual(result, expected)

    def test_get_path(self):
        expected = '/tmp'
        data = 'Temp'
        result = self.tools.getPath(data)
        self.assertEqual(result, expected)

    def test_get_path_info(self):
        expected = ('/home/mau', 'test.txt', 'test', '.txt')
        data = '/home/mau/test.txt'
        result = self.tools.getPathInfo(data)
        self.assertEqual(result, expected)

    def test_get_folder(self):
        expected = ''
        data = self.tools.getPath('Work')
        result = self.tools.getFolder(data)
        self.assertEqual(result, expected)

    def test_get_selected_files(self):
        expected = ''
        data = self.tools.getPath('Work')
        result = self.tools.getSelectedFiles(data, False, ())
        self.assertEqual(result, expected)

    def test_get_files(self):
        expected= ('/home/mau/test/file.txt', '/home/mau/test/file.py')
        data = '/home/mau/test'
        result = self.tools.getFiles(data, '')
        self.assertEqual(result, expected)

    def test_file_save_open(self):
        expected = 'LibreOffice Python'
        test_path = os.path.join(self.test_dir, 'test.txt')
        self.tools.fileSave(test_path, 'w', expected)
        result = self.tools.fileOpen(test_path, 'r', False)
        self.assertEqual(result, expected)

    def test_execute(self):
        expected = 'c203202158107906c65fb70914fe8add  /home/mau/test/file.txt\n'
        data = ('md5sum', '/home/mau/test/file.txt')
        result = self.tools.execute(data, True)
        self.assertEqual(result, expected)

    def test_set_get_config(self):
        expected = 'Python'
        result = self.tools.setConfig('best', expected)
        self.assertTrue(result)
        result = self.tools.getConfig('best')
        self.assertEqual(result, expected)

    def test_clipboard(self):
        expected = 'LibreOffice EasyDev'
        self.tools.setClipboard(expected)
        result = self.tools.getClipboard()
        self.assertEqual(result, expected)

    def test_get_epoch(self):
        now = datetime.datetime.now()
        expected = int(time.mktime(now.timetuple()))
        result = self.tools.getEpoch()
        self.assertEqual(result, expected)


    #~ def test_export_to_csv(self):
        #~ expected = """value a1,1,value c1
#~ value a2,2,value c2
#~ value a3,3,value c3
#~ """
        #~ test_path = os.path.join(self.test_dir, 'temp.csv')
        #~ data = (
            #~ ('value a1', 1, 'value c1'),
            #~ ('value a2', 2, 'value c2'),
            #~ ('value a3', 3, 'value c3'),
        #~ )
        #~ tools.export_csv(test_path, data)
        #~ with open(test_path, 'r') as f:
            #~ result = f.read()
        #~ self.assertEqual(result, expected)
#~
    #~ def test_export_to_csv_with_headers(self):
        #~ expected = """Num 1,Num 2,Num 3
#~ value a1,1,value c1
#~ value a2,2,value c2
#~ value a3,3,value c3
#~ """
        #~ test_path = os.path.join(self.test_dir, 'temp.csv')
        #~ options = (
            #~ ('write_headers', True),
            #~ ('headers', ('Num 1', 'Num 2', 'Num 3')),
        #~ )
        #~ data = (
            #~ ('VALUE 1', 'VALUE 2', 'VALUE 3'),
            #~ ('value a1', 1, 'value c1'),
            #~ ('value a2', 2, 'value c2'),
            #~ ('value a3', 3, 'value c3'),
        #~ )
        #~ tools.export_csv(test_path, data, options)
        #~ with open(test_path, 'r') as f:
            #~ result = f.read()
        #~ self.assertEqual(result, expected)
#~
    #~ def test_export_to_csv_change_delimiter(self):
        #~ expected = """VALUE 1|VALUE 2|VALUE 3
#~ value a1|1|value c1
#~ value a2|2|value c2
#~ value a3|3|value c3
#~ """
        #~ test_path = os.path.join(self.test_dir, 'temp.csv')
        #~ options = (
            #~ ('write_headers', False),
            #~ ('delimiter', '|'),
        #~ )
        #~ data = (
            #~ ('VALUE 1', 'VALUE 2', 'VALUE 3'),
            #~ ('value a1', 1, 'value c1'),
            #~ ('value a2', 2, 'value c2'),
            #~ ('value a3', 3, 'value c3'),
        #~ )
        #~ tools.export_csv(test_path, data, options)
        #~ with open(test_path, 'r') as f:
            #~ result = f.read()
        #~ self.assertEqual(result, expected)
#~
    #~ def test_get_cell_active(self):
        #~ expected = 'ScCellObj'
        #~ libo = LIBO()
        #~ doc = libo.new_doc()
        #~ cell = tools.get_cell(doc)
        #~ result = cell.getImplementationName()
        #~ del libo
        #~ self.assertEqual(result, expected)
#~
    #~ def test_get_cell_by_name(self):
        #~ expected = 'ScCellObj'
        #~ libo = LIBO()
        #~ doc = libo.new_doc()
        #~ cell = tools.get_cell(doc, 'Sheet1', 'E5')
        #~ result = cell.getImplementationName()
        #~ del libo
        #~ self.assertEqual(result, expected)
#~
    #~ def test_get_cell_by_pos(self):
        #~ expected = 'ScCellObj'
        #~ libo = LIBO()
        #~ doc = libo.new_doc()
        #~ cell = tools.get_cell(doc, 'Sheet1', (4, 4))
        #~ result = cell.getImplementationName()
        #~ del libo
        #~ self.assertEqual(result, expected)
#~
    #~ def test_select_range(self):
        #~ expected = '$Sheet1.$A$1:$E$5'
        #~ libo = LIBO()
        #~ doc = libo.new_doc()
        #~ tools.select_range(doc, '', 'A1:E5')
        #~ rango = doc.getCurrentSelection()
        #~ result = rango.AbsoluteName
        #~ del libo
        #~ self.assertEqual(result, expected)
#~
    #~ def test_get_range_selected(self):
        #~ expected = 'ScCellRangeObj'
        #~ libo = LIBO()
        #~ doc = libo.new_doc()
        #~ tools.select_range(doc, '', 'A1:E5')
        #~ range_cell = tools.get_range(doc)
        #~ result = range_cell.getImplementationName()
        #~ del libo
        #~ self.assertEqual(result, expected)
#~
    #~ def test_get_current_region(self):
        #~ expected = '$Sheet1.$A$1:$C$3'
        #~ libo = LIBO()
        #~ doc = libo.new_doc()
        #~ data = (
            #~ ('A1', 'B1', 'C1'),
            #~ ('A2', 'B2', 'C2'),
            #~ ('A3', 'B3', 'C3'),
        #~ )
        #~ doc.getSheets().getByIndex(0).getCellRangeByName('A1:C3').setDataArray(data)
        #~ cell = doc.getSheets().getByIndex(0).getCellRangeByName('A1')
        #~ rango = tools.get_current_region(cell)
        #~ result = rango.AbsoluteName
        #~ del libo
        #~ self.assertEqual(result, expected)
#~
    #~ def test_get_last_row(self):
        #~ expected = 2
        #~ libo = LIBO()
        #~ doc = libo.new_doc()
        #~ data = (
            #~ ('A1', 'B1', 'C1'),
            #~ ('A2', 'B2', 'C2'),
            #~ ('A3', 'B3', 'C3'),
        #~ )
        #~ doc.getSheets().getByIndex(0).getCellRangeByName('A1:C3').setDataArray(data)
        #~ cell = doc.getSheets().getByIndex(0).getCellRangeByName('A1')
        #~ result = tools.get_last_row(cell)
        #~ del libo
        #~ self.assertEqual(result, expected)


if __name__ == '__main__':
    unittest.main()

