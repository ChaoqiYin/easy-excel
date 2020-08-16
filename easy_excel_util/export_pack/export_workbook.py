# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import xlwt

from .export_row import ExportRow
from .export_sheet import ExportSheet
from .sheet_map import SheetMap


class ExportWorkBook(object):
    def __init__(self, file_path_or_stream, converters, style=None, title_style=None):
        self.file_path = file_path_or_stream  # 文件路径
        self.converters = converters  # 转换类
        self.style = style  # 单元格样式
        self.title_style = title_style  # 单元格头样式
        self.sheet_map = {}  # 表索引: [表名, parse_map, list_data]关系

    def sheet(self, index, data, parse_map, sheet_name=None, row_height=20, col_width=20, style=None, title_style=None):
        '''
        设置sheet对应map
        :param index: sheet索引
        :param parse_map: 解析map
        :param data: 数据
        :param sheet_name: 表名
        :param row_height: 默认行高
        :param col_width: 默认列宽
        :param style: 单元格
        :param title_style: 标题样式
        :return:
        '''
        rel_sheet_name = 'sheet_' + str(index) if sheet_name is None else sheet_name
        self.sheet_map[index] = SheetMap(rel_sheet_name, parse_map, data, row_height, col_width, style or self.style,
                                         title_style or self.title_style)
        return self

    def sort_sheet_map(self):
        keys = list(self.sheet_map.keys())
        keys.sort(reverse=False)
        new_map = {}
        for key in keys:
            new_map[key] = self.sheet_map[key]
        self.sheet_map = new_map

    def do_export(self, max_workers=None, row_del_class=None, row_validate_func=None):
        rel_row_del_class = ExportRow if row_del_class is None else row_del_class
        # 创建一个workbook 设置编码
        workbook = xlwt.Workbook(encoding='utf-8')
        # 升序排序
        self.sort_sheet_map()
        for index, sheet_map in self.sheet_map.items():
            sheet = ExportSheet(self, workbook, index, sheet_map, max_workers=max_workers,
                                row_del_class=rel_row_del_class, row_validate_func=row_validate_func)
            sheet.parse_export()
        workbook.save(self.file_path)  # file_path可能会是stream，这里因为使用了父类变量，没法改变量名
