# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .export_workbook import ExportWorkBook
from .sheet_map import SheetMap
from .export_workbook import BASE_ROW_HEIGHT, BASE_COL_WIDTH


class Middleware(object):
    '''
    导出的中间件，用于构造build.sheet.sheet...do_export的结构
    '''
    def __init__(self, converters, style=None, title_style=None, xlsx=False):
        self.__converters = converters  # 转换类
        self.__style = style  # 单元格样式
        self.__title_style = title_style  # 单元格头样式
        self.is_xlsx = xlsx

    def sheet(self, index, data, parse_map, sheet_name=None, height=BASE_ROW_HEIGHT, before=None, after=None,
              style=None, title_style=None, row_del_class=None, max_workers=None):
        '''
        设置sheet对应map
        :param index: sheet索引, 从0开始
        :param parse_map: 解析map
        :param data: 导出的数据, 用list装载
        :param sheet_name: 表名
        :param height: 默认行高
        :param before: 导入第一行前的操作
        :param after: 导入最后一行后的操作
        :param style: 该sheet单元格样式，不传则使用全局样式, 可以是fun
        :param title_style: 该sheet标题样式，不传则使用全局样式，可以是fun
        :param row_del_class: 行处理的类
        :param max_workers: 异步线程数
        :return:
        '''
        rel_sheet_name = 'sheet' + str(index + 1) if sheet_name is None else sheet_name
        rel_style = style or self.__style
        rel_title_style = title_style or self.__title_style
        # 全局样式可能为None
        sheet_map = SheetMap(index, rel_sheet_name, parse_map, data, height, rel_style, rel_title_style, row_del_class)
        return ExportWorkBook(self.__converters, sheet_map, before, after, rel_style, rel_title_style, max_workers, self.is_xlsx)