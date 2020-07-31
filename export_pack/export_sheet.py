# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from concurrent.futures import ThreadPoolExecutor, wait

from ..base_sheet import BaseSheet


class ExportSheet(BaseSheet):

    def __init__(self, excel_workbook, wb, row_del_class, row_validate_func, sheet_no, sheet_info):
        '''
        init
        :param excel_workbook: BaseWorkbook子类实例
        :param wb: workbook待写入文件
        :param row_del_class: 解析处理row的类
        :param row_validate_func: 行处理方法
        :param sheet_no: 表格索引
        :param sheet_info: sheet_name, parse_map, data 待导出的信息
        '''
        super().__init__(excel_workbook, row_del_class, row_validate_func, sheet_no)
        self.sheet_info = sheet_info
        self.wb = wb
        self.sheet_name = sheet_info.sheet_name
        self.parse_map = sheet_info.parse_map
        self.list_data = sheet_info.list_data
        self.title_style = sheet_info.title_style
        self.work_sheet = None

    def sync_parse(self):
        '''
        同步方式进行解析
        :return:
        '''
        row_num = 0
        # 后期可能会添加自定义抬头
        self.add_title(row_num)
        for row_data in self.list_data:
            row_num += 1
            self.add_row(row_num, row_data)

    def thread_parse(self):
        '''
        异步模式进行解析
        :return:
        '''
        pass

    def set_row_height(self, row_num):
        '''
        设置行高
        :param row_num:
        :return:
        '''
        self.work_sheet.row(row_num).height_mismatch = True
        self.work_sheet.row(row_num).height = 40 * 20  # 20为基准数

    def add_title(self, row_num):
        '''
        添加sheet标题栏
        :return:
        '''
        self.row_del_class(self, row_num, []).write_title()
        self.set_row_height(row_num)

    def add_row(self, row_num, row_data):
        self.row_del_class(self, row_num, row_data).write_row()
        self.set_row_height(row_num)

    def parse_export(self):
        self.work_sheet = self.wb.add_sheet(self.sheet_name)
        if self.excel_workbook.sync is True:
            self.sync_parse()
        else:
            self.thread_parse()
