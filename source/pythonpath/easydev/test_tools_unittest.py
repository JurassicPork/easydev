#!/usr/bin/env python3
# coding: utf-8

import uno
import unittest
import tempfile
import os
import tools


class TestTools(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        pass

    def test_render(self):
        expected = 'Nombre: Teresa\nPaís: Portugal'
        template = 'Nombre: $name\nPaís: $country'
        data = (('name', 'Teresa'), ('country', 'Portugal'))
        result = tools.render(template, data)
        self.assertEqual(result, expected)

    def test_render_number(self):
        expected = 'Age: 40\nPhone: 12345678'
        template = 'Age: $age\nPhone: $phone'
        data = (('age', 40), ('phone', 12345678))
        result = tools.render(template, data)
        self.assertEqual(result, expected)

    def test_file_open(self):
        expected = 'LibreOffice Python'
        test_path = os.path.join(self.test_dir, 'test.txt')
        with open(test_path, 'w') as f:
            f.write(expected)
        result = tools.file_open(test_path)
        self.assertEqual(result, expected)

    def test_file_open_binary(self):
        expected = b'LibreOffice Python'
        test_path = os.path.join(self.test_dir, 'test.txt')
        with open(test_path, 'wb') as f:
            f.write(expected)
        result = tools.file_open(test_path, 'rb')
        self.assertEqual(result, expected)

    def test_format_string(self):
        expected = 'Number: one'
        template = 'Number: {}'
        data = 'one'
        result = tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_number(self):
        expected = '123'
        template = '{:d}'
        data = 123
        result = tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_list(self):
        expected = 'Numbers: one and two'
        template = 'Numbers: {} and {}'
        data = ('one', 'two')
        result = tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_list_pos(self):
        expected = 'Numbers: two and one'
        template = 'Numbers: {1} and {0}'
        data = ('one', 'two')
        result = tools.format(template, data)
        self.assertEqual(result, expected)

    def test_format_dict(self):
        expected = 'Numbers: one and two'
        template = 'Numbers: {n1} and {n2}'
        data = (('n1', 'one'), ('n2', 'two'))
        result = tools.format(template, data)
        self.assertEqual(result, expected)

    def test_export_to_csv(self):
        expected = """value a1,1,value c1
value a2,2,value c2
value a3,3,value c3
"""
        test_path = os.path.join(self.test_dir, 'temp.csv')
        data = (
            ('value a1', 1, 'value c1'),
            ('value a2', 2, 'value c2'),
            ('value a3', 3, 'value c3'),
        )
        tools.export_csv(test_path, data)
        with open(test_path, 'r') as f:
            result = f.read()
        self.assertEqual(result, expected)

    def test_export_to_csv_with_headers(self):
        expected = """Num 1,Num 2,Num 3
value a1,1,value c1
value a2,2,value c2
value a3,3,value c3
"""
        test_path = os.path.join(self.test_dir, 'temp.csv')
        options = (
            ('write_headers', True),
            ('headers', ('Num 1', 'Num 2', 'Num 3')),
        )
        data = (
            ('VALUE 1', 'VALUE 2', 'VALUE 3'),
            ('value a1', 1, 'value c1'),
            ('value a2', 2, 'value c2'),
            ('value a3', 3, 'value c3'),
        )
        tools.export_csv(test_path, data, options)
        with open(test_path, 'r') as f:
            result = f.read()
        self.assertEqual(result, expected)

    def test_export_to_csv_change_delimiter(self):
        expected = """VALUE 1|VALUE 2|VALUE 3
value a1|1|value c1
value a2|2|value c2
value a3|3|value c3
"""
        test_path = os.path.join(self.test_dir, 'temp.csv')
        options = (
            ('write_headers', False),
            ('delimiter', '|'),
        )
        data = (
            ('VALUE 1', 'VALUE 2', 'VALUE 3'),
            ('value a1', 1, 'value c1'),
            ('value a2', 2, 'value c2'),
            ('value a3', 3, 'value c3'),
        )
        tools.export_csv(test_path, data, options)
        with open(test_path, 'r') as f:
            result = f.read()
        self.assertEqual(result, expected)


if __name__ == '__main__':
    unittest.main()
