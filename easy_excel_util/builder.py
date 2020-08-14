# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .import_pack.import_workbook import ImportWorkbook
from .export_pack.export_workbook import ExportWorkBook
from .utils import get_converters_key

# 导入/导出时的数据全局转换方法，key对应ctype_text中的类型
_converters = {}


class Builder(object):
    '''
    构造类，主要用于设置导入/导出时的数据转换方法，下面的key对应ctype_text中的类型
    '''
    @classmethod
    def add_import_converter(cls, converter_key, func):
        '''
        设置导入转换类
        :param converter_key: 转换类型key
        :param func: 转换方法，接收一个value参数，值为单元格数据
        :return:
        '''
        if converter_key not in range(0, 7):
            raise Exception("converter_key must in the [0, 1, 2, 3, 4, 5, 6] keys")
        # 检查是否已经设置过这个key的转换方法，提醒
        key = get_converters_key(converter_key, True)
        if _converters.get(key, None) is not None:
            raise Exception("already set import_converter with key: %s" % key)
        _converters[key] = func
        return cls

    @classmethod
    def add_export_converter(cls, data_type_class, func):
        '''
        设置导出转换类
        :param data_type_class: 转换类型的类名: type(x)
        :param func: 转换方法，接收一个value参数，值为待转换数据
        :return:
        '''
        # 检查是否已经设置过这个key的转换方法，提醒
        key = get_converters_key(data_type_class, False)
        if _converters.get(key, None) is not None:
            raise Exception("already set export_converter with data_type_class: %s" % key)
        _converters[key] = func
        return cls

    @staticmethod
    def build_import(file, parse_map, sheet_no=0, start_row_num=0, end_row_num=None):
        '''
        生成导入实例
        :param file: 文件路径或者可read()的文件
        :param parse_map: 解析的字典
        :param sheet_no: 解析的表格索引
        :param start_row_num: 从第几行开始解析
        :param end_row_num: 到第几行结束
        :return:
        '''
        return ImportWorkbook(file, parse_map, _converters.copy(), sheet_no, start_row_num, end_row_num)

    @staticmethod
    def build_export(file_path_or_stream):
        return ExportWorkBook(file_path_or_stream, _converters.copy())
