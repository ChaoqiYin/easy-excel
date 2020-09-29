# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import xlwt

from .base import Base, type_factory


# 设置边框
borders = xlwt.Borders()
# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1

# 单元格默认样式, 默认是10号字体
DEFAULT_STYLE = xlwt.XFStyle()
al = xlwt.Alignment()
al.wrap = 1  # 自动换行
al.horz = 0x02  # 设置水平居中
al.vert = 0x01  # 设置垂直居中
DEFAULT_STYLE.alignment = al
font = xlwt.Font()  # 为样式创建字体
# 字体大小，14为字号，20为衡量单位
font.name = 'Arial'
font.height = 10 * 20
DEFAULT_STYLE.font = font
DEFAULT_STYLE.borders = borders

# 单元格默认title样式
DEFAULT_TITLE_STYLE = xlwt.XFStyle()
al2 = xlwt.Alignment()
al2.wrap = 1
al2.horz = 0x02  # 设置水平居中
al2.vert = 0x01  # 设置垂直居中
DEFAULT_TITLE_STYLE.alignment = al2
title_font = xlwt.Font()  # 为样式创建字体
# 字体大小，14为字号，20为衡量单位
title_font.name = 'Arial'
title_font.height = 14 * 20
title_font.bold = True  # 黑体
DEFAULT_TITLE_STYLE.font = title_font  # 设定样式
DEFAULT_TITLE_STYLE.borders = borders


class Xls(Base):
    def __init__(self):
        super().__init__()
        self.wb = xlwt.Workbook(encoding='utf-8')

    def save(self, file_path_or_stream):
        if isinstance(file_path_or_stream, str) and file_path_or_stream.find('.xls') < 0:
            file_path_or_stream += '.xls'
        self.wb.save(file_path_or_stream)

    def get_sheet(self, sheet_name):
        return self._sheet_map.get(sheet_name)

    def add_sheet(self, sheet_name):
        self._sheet_map[sheet_name] = self.wb.add_sheet(sheet_name, cell_overwrite_ok=True)

    def merge(self, sheet_name, start_row, end_row, start_col, end_col, style):
        sheet = self.get_sheet(sheet_name)
        sheet.merge(start_row, end_row, start_col, end_col, self._get_style(style, False))

    def set_row_height(self, sheet_name, row_num, row_height):
        sheet = self.get_sheet(sheet_name)
        sheet.row(row_num).height_mismatch = True
        sheet.row(row_num).height = 20 * row_height  # 20为基准数

    def set_col_width(self, sheet_name, col_num, col_width):
        sheet = self.get_sheet(sheet_name)
        sheet.col(col_num).width = 256 * col_width  # 256为基准数

    @staticmethod
    def _get_style(style, is_title):
        # 有样式直接使用该样式，否则使用默认样式，区分是title样式或普通样式
        if style is not None:
            return style
        return DEFAULT_STYLE if is_title is False else DEFAULT_TITLE_STYLE

    def write(self, sheet_name, row_num, col_num, value, style, is_title=False):
        sheet = self.get_sheet(sheet_name)
        sheet.write(row_num, col_num, value, self._get_style(style, is_title))


type_factory['xls'] = Xls
