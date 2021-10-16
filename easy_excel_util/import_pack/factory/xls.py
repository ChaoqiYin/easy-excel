# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from xlrd import xldate_as_datetime

from .base import Base, Cell, type_factory


class Xls(Base):
    def __init__(self, sheet, sheet_no):
        super().__init__(sheet, sheet_no)

    @property
    def merged_cells(self):
        '''
        返回合并单元格的列表[(start_row, end_row, start_col, end_col)...], tuple前闭后开
        :return:
        '''
        return self.sheet.merged_cells

    @property
    def nrows(self):
        return self.sheet.nrows

    @property
    def ncols(self):
        return self.sheet.ncols

    def row(self, row_num):
        return self.sheet.row(row_num)

    def cell(self, row_num, col_num):
        cell = self.sheet.cell(row_num, col_num)
        return Cell(cell.ctype, cell.value)

    @staticmethod
    def del_datetime(value):
        '''
        cell时间内容转换为datetime
        :param value: cell单元格数据，时间类型
        :return:
        '''
        return xldate_as_datetime(value, 0)


type_factory['xls'] = Xls
