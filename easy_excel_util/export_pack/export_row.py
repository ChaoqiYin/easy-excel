# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .export_cell import ExportCell


class ExportRow(object):
    def __init__(self, export_sheet, row_num, row_data):
        '''
        init
        :param export_sheet: ExportSheet实例
        :param row_num: 当前行数，从0开始
        :param row_data: 当前行数据
        '''
        self.export_sheet = export_sheet
        self.row_num = row_num
        self.row_data = row_data

    @property
    def sheet(self):
        return self.export_sheet.work_sheet

    @property
    def parse_map(self):
        '''
        获取解析字段的map映射
        :return:
        '''
        return self.export_sheet.sheet_map.parse_map

    def set_row_height(self):
        '''
        设置行高
        :param row_num:
        :return:
        '''
        self.sheet.row(self.row_num).height_mismatch = True
        self.sheet.row(self.row_num).height = 40 * self.export_sheet.sheet_map.row_height  # 20为基准数

    def write_title(self):
        '''
        写入title数据
        :return:
        '''
        self.set_row_height()
        for export_field_name, export_field in self.parse_map.items():
            ExportCell(self, export_field.name, export_field).write_title_cell()

    def write_row(self):
        '''
        写入行数据
        :return:
        '''
        self.set_row_height()
        for export_field_name, export_field in self.parse_map.items():
            # 判断是否有这个key
            if self.row_data.get(export_field_name, None) is not None:
                ExportCell(self, self.row_data[export_field_name], export_field).write_cell()