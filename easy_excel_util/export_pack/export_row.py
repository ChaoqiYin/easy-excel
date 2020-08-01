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

    def get_parse_map(self):
        '''
        获取解析字段的map映射
        :return:
        '''
        return self.export_sheet.parse_map

    def write_title(self):
        '''
        写入title数据
        :return:
        '''
        for export_field_name, export_field in self.get_parse_map().items():
            ExportCell(self, export_field.col_name or export_field.name, export_field).write_title_cell()

    def write_row(self):
        '''
        写入行数据
        :return:
        '''
        for export_field_name, export_field in self.get_parse_map().items():
            # 判断是否有这个key
            if self.row_data.get(export_field_name, None) is not None:
                ExportCell(self, self.row_data[export_field_name], export_field).write_cell()