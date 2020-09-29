# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .import_pack.import_workbook import ImportWorkbook
from .export_pack.export_middleware import Middleware


class Builder(object):

    # 导入/导出时的数据全局转换方法，key对应ctype_text中的类型
    import_converters = {}
    export_converters = {}
    # 导出通用样式
    _export_style = None
    # 导出title样式，即第一行样式
    _export_title_style = None

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
        if cls.import_converters.get(converter_key, None) is not None:
            raise Exception("already set import_converter with key: %s" % converter_key)
        cls.import_converters[converter_key] = func
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
        if cls.export_converters.get(data_type_class, None) is not None:
            raise Exception("already set export_converter with data_type_class: %s" % key)
        cls.export_converters[data_type_class] = func
        return cls

    @classmethod
    def add_export_style(cls, style, title_style=None):
        cls._export_style = style
        cls._export_title_style = title_style
        return cls

    @classmethod
    def build_import(cls, file):
        '''
        生成导入实例
        :param file: 文件，可.read()对象
        :return:
        '''
        return ImportWorkbook(file, cls.import_converters.copy())

    @classmethod
    def build_export(cls, xlsx=False):
        '''
        :return:
        '''
        return Middleware(cls.export_converters.copy(), cls._export_style, cls._export_title_style, xlsx)
