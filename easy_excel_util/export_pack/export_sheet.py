# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from concurrent.futures import ThreadPoolExecutor, wait

from .export_row import ExportRow


class ExportSheet(object):

    def __init__(self, excel_workbook, workbook, sheet_no, sheet_map, max_workers, row_del_class, row_validate_func):
        '''
        init
        :param excel_workbook: BaseWorkbook子类实例
        :param workbook: workbook待写入文件
        :param sheet_no: 表格索引
        :param sheet_map: sheet_map实例
        :param max_workers: max_workers
        :param row_del_class: 解析处理row的类
        :param row_validate_func: 行处理方法
        '''
        self.excel_workbook = excel_workbook
        self.workbook = workbook
        self.work_sheet = workbook.add_sheet(sheet_map.sheet_name)
        self.sheet_no = sheet_no
        self.sheet_map = sheet_map
        self.max_workers = max_workers
        self.row_del_class = row_del_class
        self.row_validate_func = row_validate_func  # 自定义的行处理方法，未自定义时为None


    def sync_parse(self):
        '''
        同步方式进行解析
        :return:
        '''
        row_num = 0
        # 后期可能会添加自定义抬头
        self.add_title(row_num)
        for row_data in self.sheet_map.list_data:
            row_num += 1
            self.add_row(row_num, row_data)

    def thread_parse(self):
        '''
        异步模式进行解析
        :return:
        '''
        def work_func(that, rn, data):
            that.add_row(rn, data)

        work_list = []
        # 创建线程池
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            row_num = 0
            # 后期可能会添加自定义抬头
            self.add_title(row_num)
            for row_data in self.sheet_map.list_data:
                row_num += 1
                future = executor.submit(work_func, self, row_num, row_data)
                work_list.append(future)
            # 等待完成
            wait(work_list)

    def set_row_height(self, row_num):
        '''
        设置行高
        :param row_num:
        :return:
        '''
        self.work_sheet.row(row_num).height_mismatch = True
        self.work_sheet.row(row_num).height = 40 * self.sheet_map.row_height  # 20为基准数

    def add_title(self, row_num):
        '''
        添加sheet标题栏
        :return:
        '''
        self.set_row_height(row_num)
        ExportRow(self, row_num, []).write_title()

    def add_row(self, row_num, row_data):
        self.set_row_height(row_num)
        self.row_del_class(self, row_num, row_data).write_row()

    def parse_export(self):
        if self.max_workers is None:
            self.sync_parse()
        else:
            self.thread_parse()
