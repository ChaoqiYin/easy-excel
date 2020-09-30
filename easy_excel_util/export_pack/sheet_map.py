# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


class SheetMap(object):
    def __init__(self, sheet_name, parse_map, data, height, style, title_style, row_del_class):
        self.sheet_name = sheet_name  # 表名
        self.parse_map = parse_map  # parse_map
        self.list_data = data  # 待解析数据
        self.row_height = height
        self.style = style  # 单元格样式
        self.title_style = title_style  # excel头样式
        self.row_del_class = row_del_class  # 行处理类
