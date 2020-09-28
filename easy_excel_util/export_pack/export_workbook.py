# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import xlwt

from .export_row import ExportRow
from .export_sheet import ExportSheet
from .sheet_map import SheetMap


class ExportWorkBook(object):
    def __init__(self, file_path_or_stream, converters, sheet_map, style=None, title_style=None, max_workers=None):
        self.file_path = file_path_or_stream  # 文件路径
        self.converters = converters  # 转换类
        self.style = style  # 单元格样式
        self.title_style = title_style  # 单元格头样式
        self.max_workers = max_workers
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.__start_parse(sheet_map)

    def sheet(self, index, data, parse_map, sheet_name=None, row_height=40, col_width=250, style=None, title_style=None, row_del_class=None):
        '''
        设置sheet对应map
        :param index: sheet索引, 从0开始
        :param parse_map: 解析map
        :param data: 导出的数据, 用list装载
        :param sheet_name: 表名
        :param row_height: 默认行高
        :param col_width: 默认列宽
        :param style: 该sheet单元格样式，不传则使用全局样式
        :param title_style: 该sheet标题样式，不传则使用全局样式
        :param row_del_class: 行处理类
        :return:
        '''
        rel_sheet_name = 'sheet_' + str(index + 1) if sheet_name is None else sheet_name
        rel_style = style or self.style
        rel_title_style = title_style or self.title_style
        # 全局样式可能为None
        sheet_map = dict(index=SheetMap(rel_sheet_name, parse_map, data, row_height, col_width, rel_style, rel_title_style, row_del_class))
        self.__start_parse(sheet_map)
        return self

    def __start_parse(self, sheet_map):
        # 升序排序
        for index, sm in sheet_map.items():
            rel_row_del_class = sm.row_del_class or ExportRow
            sheet = ExportSheet(self, self.workbook, index, sm, max_workers=self.max_workers, row_del_class=rel_row_del_class)
            sheet.parse_export()

    def do_export(self):
        if isinstance(self.file_path, str) and self.file_path.find('.xls') < 0:
            self.file_path += '.xls'
        self.workbook.save(self.file_path)  # file_path可能会是stream，这里因为使用了父类变量，没法改变量名
