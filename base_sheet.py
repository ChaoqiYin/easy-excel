# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


class BaseSheet(object):
    def __init__(self, excel_workbook, row_del_class, row_validate_func, sheet_no):
        '''
        init
        :param excel_workbook: BaseWorkbook子类实例
        :param row_del_class: 解析处理row的类
        :param row_validate_func: 行处理方法
        :param sheet_no: sheet索引
        '''
        self.excel_workbook = excel_workbook
        self.row_del_class = row_del_class
        self.row_validate_func = row_validate_func  # 自定义的行处理方法，未自定义时为None
        self.sheet_no = sheet_no
