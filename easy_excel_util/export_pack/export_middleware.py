# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .export_workbook import ExportWorkBook
from .sheet_map import SheetMap


class Middleware(object):
    '''
    导出的中间件，用于构造build.sheet.sheet...do_export的结构
    '''
    def __init__(self, file_path_or_stream, converters, style=None, title_style=None, max_workers=None):
        self.__file_path = file_path_or_stream  # 文件路径
        self.__converters = converters  # 转换类
        self.__style = style  # 单元格样式
        self.__title_style = title_style  # 单元格头样式
        self.__max_workers = max_workers

    def sheet(self, index, data, parse_map, sheet_name=None, row_height=40, col_width=250, style=None,
              title_style=None, row_del_class=None):
        '''
        设置sheet对应map
        :param index: sheet索引, 从0开始
        :param parse_map: 解析map
        :param data: 导出的数据, 用list装载
        :param sheet_name: 表名
        :param row_height: 默认行高
        :param col_width: 默认列宽
        :param style: 该sheet单元格样式，不传则使用全局样式, 可以是fun
        :param title_style: 该sheet标题样式，不传则使用全局样式，可以是fun
        :param row_del_class: 行处理的类
        :return:
        '''
        rel_sheet_name = 'sheet_' + str(index + 1) if sheet_name is None else sheet_name
        rel_style = style or self.__style
        rel_title_style = title_style or self.__title_style
        # 全局样式可能为None
        sheet_map = dict(index=SheetMap(rel_sheet_name, parse_map, data, row_height, col_width, rel_style, rel_title_style, row_del_class))
        return ExportWorkBook(self.__file_path, self.__converters, sheet_map, self.__style, self.__title_style)