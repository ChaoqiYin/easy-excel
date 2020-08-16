# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


class SheetMap(object):
    def __init__(self, sheet_name, parse_map, list_data, row_height, col_width, style, title_style):
        self.sheet_name = sheet_name  # 表名
        self.parse_map = parse_map  # parse_map
        self.list_data = list_data  # 待解析数据
        self.row_height = row_height
        self.col_width = col_width
        self.style = style  # 单元格样式
        self.title_style = title_style  # excel头样式
