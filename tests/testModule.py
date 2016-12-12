from unittest import TestCase
import os
import sys
import xlrd
from .base import from_this_dir

class TestWriteAccess(TestCase):
    def setUp(self):
        book = xlrd.open_workbook(from_this_dir('write-access-test.xls'))
        self.sheet = book.sheet_by_index(0)

    def get_value(self,col,row):
        return ascii(self.sheet.col_values(col)[row])

    def test_able_to_read_first_position(self):
        self.assertEqual(self.get_value(0,1),'Registered Nurse SP II')