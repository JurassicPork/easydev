#!/usr/bin/env python3
# coding: utf-8

import unittest
import tools


class TestTools(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_render(self):
        template = 'Nombre: $name\nPaís: $country'
        data = (('name', 'Teresa'), ('country', 'Portugal'))
        result = tools.render(template, data)
        self.assertEqual('Nombre: Teresa\nPaís: Portugal', result)

    def test_render_number(self):
        template = 'Age: $age\nPhone: $phone'
        data = (('age', 40), ('phone', 12345678))
        result = tools.render(template, data)
        self.assertEqual('Age: 40\nPhone: 12345678', result)

if __name__ == '__main__':
    unittest.main()
