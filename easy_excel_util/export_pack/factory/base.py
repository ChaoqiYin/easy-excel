# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
type_factory = {}


class Base(object):
    def __init__(self):
        self._sheet_map = {}  # {'sheet_name': sheet}

    def save(self, file_path_or_stream):
        '''
        需要实现存储到路径和存储到stream
        :param file_path_or_stream:
        :return:
        '''
        return

    def get_sheet(self, sheet_name):
        return

    def add_sheet(self, sheet_name):
        return

    def merge(self, sheet_name, start_row, end_row, start_col, end_col, style):
        return

    def set_row_height(self, sheet_name, row_num, row_height):
        return

    def set_col_width(self, sheet_name, col_num, col_width):
        return

    def write(self, sheet_name, row_num, col_num, value, style, is_title=False):
        return