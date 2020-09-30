# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import datetime
import time
from inspect import isfunction


class ExportCell(object):

    def __init__(self, export_row, cell_data, export_field):
        '''
        init
        :param export_row: ExportRow实例
        :param cell_data: 单元格内容
        :param export_field: 导出字段设置
        '''
        self.export_row = export_row
        self.cell_data = cell_data
        self.export_field = export_field

    @property
    def excel(self):
        return self.export_row.excel

    @property
    def sheet_name(self):
        return self.export_row.sheet_name

    def get_converter(self, cell_data):
        '''
        获取当前应该使用的转换方法
        :param cell_data: 当前cell的数据
        :return:
        '''
        # 如果字段本身设置了转换方法，则优先使用此转换方法
        if self.export_field.converter is not None:
            return self.export_field.converter
        # 匹配excel_workbook的转换方法
        return self.export_row.export_sheet.excel_workbook.converters.get(type(cell_data), None)

    def converter_del_value(self, cell_data):
        '''
        根据转换方法处理数据
        :param cell_data: 当前cell的数据
        :return:
        '''
        converter = self.get_converter(cell_data)
        if converter is None:
            return cell_data
        else:
            return converter(cell_data)

    def format_str_to_datetime(self, value):
        '''
        根据字段format转换为str
        :param value: datetime或者time
        :return:
        '''
        if self.export_field.datetime_format is None:
            try:
                return value.strftime('%Y-%m-%d %H:%M:%S')  # 默认时间格式
            except Exception:
                return time.strftime("%Y-%m-%d %H:%M:%S", value)
        return value.strftime(self.export_field.datetime_format)

    def write_title_cell(self):
        '''
        写入title单元格数据
        :return:
        '''
        self.excel.write(self.sheet_name, self.export_row.row_num, self.export_field.index, str(self.cell_data or ''),
                         self.export_row.export_sheet.sheet_map.title_style, is_title=True)

    def get_style(self, value):
        '''
        获取普通单元格样式
        :param value: 转换后的值
        :param is_title: 是否是title
        :return:
        '''
        if self.export_field.style is not None:
            if isfunction(self.export_field.style) is True:  # 如果style是一个方法的话
                return self.export_field.style(self.export_row.row_num, self.cell_data, value)  # 根据row_num、cell_data、转换值计算style
            else:
                return self.export_field.style
        return self.export_row.export_sheet.sheet_map.style

    def write_cell(self):
        '''
        写入正常单元格数据
        :return:
        '''
        value = self.cell_data
        value = self.converter_del_value(value)  # 首先转换
        if type(value) == datetime.datetime or type(value) == time.struct_time:
            value = self.format_str_to_datetime(value)  # 然后格式化
        style = self.get_style(value)
        self.excel.write(self.sheet_name, self.export_row.row_num, self.export_field.index, str(value), style, is_title=False)
        return dict(value=str(value), style=style)
