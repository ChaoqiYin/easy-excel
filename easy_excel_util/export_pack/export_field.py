# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from ..base_field import BaseField
from .export_workbook import BASE_COL_WIDTH


class ExportField(BaseField):
    def __init__(self, index, datetime_format=None, col_name=None, width=BASE_COL_WIDTH, converter=None, style=None, merge_same=False):
        '''
        init
        :param index:
        :param datetime_format:
        :param col_name:
        :param converter:
        :param style: 字段对应单元格样式
        :param merge_same: 是否合并相同数据的竖向列单元格
        '''
        super().__init__(index, datetime_format, col_name, converter)
        self.style = style
        self.merge_same = merge_same
        self.width = width
