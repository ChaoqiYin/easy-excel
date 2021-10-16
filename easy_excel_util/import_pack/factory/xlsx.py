# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .base import Base, Cell, type_factory


class Xlsx(Base):
    def __init__(self, sheet, sheet_no):
        super().__init__(sheet, sheet_no)

    @property
    def merged_cells(self):
        '''
        返回合并单元格的列表[(start_row, end_row, start_col, end_col)...], tuple前闭后开
        :return:
        '''
        merged_list = []
        for m in self.sheet.merged_cell_ranges:
            merged_list.append((
                m.min_row - 1, m.max_row, m.min_col - 1, m.max_col
            ))
        return merged_list

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

    @staticmethod
    def convert_cell(cell):
        '''
        转换type的值，因为最开始是用xlrd写的，要转换为对应的类型值
        ctype类型关系
            0: 'empty',
            1: 'text',
            2: 'number',
            3: 'xldate',
            4: 'bool',
            5: 'error',
            6: 'blank'
        :param cell: cell单元格在openpyxl中的值
        '''
        type_value = cell.data_type
        value = cell.value
        if type_value == 'n' and value is None:
            return Cell(0, value)
        elif type_value == 's':  # 字符串格式
            return Cell(1, value)
        elif type_value == 'n' and value is not None:  # 数字格式
            return Cell(2, value)
        elif type_value == 'd':  # 时间格式
            return Cell(3, value)
        elif type_value == 'b':  # bool格式
            return Cell(4, 0 if value is False else True)
        elif type_value == 'e':  # 错误类型
            return Cell(5, value)
        else:
            return Cell(6, None)

    def cell(self, row_num, col_num):
        cell = self.sheet.cell(row=row_num + 1, column=col_num + 1)
        return self.convert_cell(cell)

    @staticmethod
    def del_datetime(value):
        '''
        cell时间内容转换为datetime, xlsx读取出来本身就是dateTime，不需要转换
        :param value: cell单元格数据，时间类型
        :return:
        '''
        return value


type_factory['xlsx'] = Xlsx
