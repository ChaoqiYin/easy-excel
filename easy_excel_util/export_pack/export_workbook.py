# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .export_row import ExportRow
from .export_sheet import ExportSheet
from .sheet_map import SheetMap
from .factory import type_factory

BASE_ROW_HEIGHT = 40
BASE_COL_WIDTH = 20


class ExportWorkBook(object):
    def __init__(self, converters, sheet_map, before=None, after=None, style=None,
                 title_style=None, max_workers=None, xlsx=False):
        self.converters = converters  # 转换类
        self.style = style  # 单元格样式
        self.title_style = title_style  # 单元格头样式
        self.is_xlsx = xlsx
        # 根据是否是xlsx生成不同的workbook
        self.excel = type_factory.get('xls' if xlsx is False else 'xlsx')()
        self.__start_parse(sheet_map, before, after, max_workers)

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
        :param after: 导入完成后的操作
        :param style: 该sheet单元格样式，不传则使用全局样式
        :param title_style: 该sheet标题样式，不传则使用全局样式
        :param row_del_class: 行处理类
        :param max_workers: 异步线程数
        :return:
        '''
        rel_sheet_name = 'sheet_' + str(index + 1) if sheet_name is None else sheet_name
        rel_style = style or self.style
        rel_title_style = title_style or self.title_style
        # 全局样式可能为None
        sheet_map = dict(index=SheetMap(rel_sheet_name, parse_map, data, height, rel_style, rel_title_style, row_del_class))
        self.__start_parse(sheet_map, before, after, max_workers)
        return self

    def __start_parse(self, sheet_map, before, after, max_workers):
        # 升序排序
        for index, sm in sheet_map.items():
            rel_row_del_class = sm.row_del_class or ExportRow
            ExportSheet(self, self.excel, index, sm, max_workers=max_workers, row_del_class=rel_row_del_class).parse_export(before, after)

    def do_export(self, file_path_or_stream):
        self.excel.save(file_path_or_stream)  # file_path可能会是stream
