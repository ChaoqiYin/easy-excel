# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


from .base import Base, Cell, type_factory


class Xlsx(Base):
    def __init__(self, workbook, sheet_no):
        super().__init__(workbook, sheet_no)
        self.sheet = workbook.worksheets[sheet_no]

    @property
    def merged_cells(self):
        '''
        返回合并单元格的列表[(start_row, end_row, start_col, end_col)...]
        :return:
        '''
        return []

    @property
    def nrows(self):
        return self.sheet.max_row

    @property
    def ncols(self):
        return self.sheet.max_column

    def row(self, row_num):
        rowdata = []
        for i in range(0, self.ncols):
            cell = self.sheet.cell(row=row_num + 1, column=i + 1)
            rowdata.append(cell)
        return rowdata

    def cell(self, row_num, col_num):
        cell = self.sheet.cell(row=row_num + 1, column=col_num + 1)
        return Cell(None, cell.value)

    @staticmethod
    def del_datetime(value):
        '''
        cell时间内容转换为datetime
        :param value: cell单元格数据，时间类型
        :return:
        '''
        return ''


type_factory['xlsx'] = Xlsx
