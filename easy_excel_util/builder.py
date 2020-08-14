# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .import_pack.import_workbook import ImportWorkbook
from .export_pack.export_workbook import ExportWorkBook
from .utils import get_converters_key


class Builder(object):
    '''
    构造类，主要用于设置导入/导出时的数据转换方法，下面的key对应ctype_text中的类型
    '''

    def __init__(self):
        '''
        初始化
        '''
        # 转换方法对应
        self.converters = {}
        return

    def add_import_converter(self, converter_key, func):
        '''
        设置导入转换类
        :param converter_key: 转换类型key
        :param func: 转换方法
        :return:
        '''
        if converter_key not in range(0, 7):
            raise Exception("converter_key must in the [0, 1, 2, 3, 4, 5, 6] keys")
        self.converters[get_converters_key(converter_key, True)] = func
        return self

    def add_export_converter(self, data_type_class, func):
        '''
        设置导出转换类
        :param data_type_class: 转换类型的类名: type(x)
        :param func: 转换方法
        :return:
        '''
        self.converters[get_converters_key(data_type_class, False)] = func
        return self

    def build_import(self, file, parse_map, sheet_no=0, start_row_num=0, end_row_num=None):
        '''
        生成导入实例
        :param file: 文件路径或者可read()的文件
        :param parse_map: 解析的字典
        :param sheet_no: 解析的表格索引
        :param start_row_num: 从第几行开始解析
        :param end_row_num: 到第几行结束
        :return:
        '''
        return ImportWorkbook(file, parse_map, self.converters.copy(), sheet_no, start_row_num, end_row_num)

    def build_export(self, file_path_or_stream):
        return ExportWorkBook(file_path_or_stream, self.converters.copy())
