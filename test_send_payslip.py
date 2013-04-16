#!/usr/bin/env python2.7
# encoding: utf-8

"""
test_send_payslip.py

Created by 4aiur on 2012-02-20.
Copyright (c) 2012 4aiur. All rights reserved.
"""


import unittest
import os
import sys
from pprint import pprint
cwd = os.getcwd()
app_path = os.path.abspath(os.path.join(cwd, "../"))
sys.path.append(app_path)
import send_payslip


class TestExcel(unittest.TestCase):

    def setUp(self):
        self.main_config = send_payslip.MainConfig(os.path.join("conf", "main.cfg"))
        self.main_config.get_configs()
        self.main_config.sheet_names = "CN,EN"
        self.excel = send_payslip.Excel("example.xls", self.main_config)
        return

    def tearDown(self):
        pass

    def test_get_sheets(self):
        sheets = self.excel.get_sheets()
        for sheet in sheets:
            print sheet.name

    def test_get_column_names_line_number(self):
        sheet = self.excel.get_sheets()[0]
        self.excel.key_column = set([u"员工姓名", u"邮箱"])
        self.assertEqual(self.excel.get_column_names_line_number(sheet), 0)
        sheet = self.excel.get_sheets()[1]
        self.excel.key_column = set([u"Compellation", u"Email"])
        self.assertEqual(self.excel.get_column_names_line_number(sheet), 0)

    def test_get_columns(self):
        sheet = self.excel.get_sheets()[0]
        self.excel.get_columns(sheet)

    def test_get_employees(self):
        sheet = self.excel.get_sheets()[0]
        columns = self.excel.get_columns(sheet)
        self.excel.get_employees(sheet, columns)


if __name__ == "__main__":
    unittest.main()
    #unittest.TextTestRunner(verbosity=2).run(suite)
