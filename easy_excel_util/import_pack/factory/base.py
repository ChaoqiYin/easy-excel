# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
type_factory = {}


class Base(object):
    def __init__(self, sheet, sheet_no):
        self.sheet = sheet
        self.sheet_no = sheet_no

    @property
    def nrows(self):
        return

    @property
    def ncols(self):
        return

    def cell(self, row_num, col_num):
        return

    @staticmethod
    def del_datetime(value):
        return


class Cell(object):
    def __init__(self, cell_type, cell_value):
        self.ctype = cell_type
        self.value = cell_value