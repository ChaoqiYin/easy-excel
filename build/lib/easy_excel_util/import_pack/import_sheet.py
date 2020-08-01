# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from concurrent.futures import ThreadPoolExecutor, wait

from ..utils import sort_dict_data, sort_dict_list_data
from .result import Result
from ..base_sheet import BaseSheet


class ImportSheet(BaseSheet):
    def __init__(self, excel_workbook, sheet, sheet_no, start_row_num,
                 row_del_class, error_message_prefix, row_validate_func):
        '''
        init
        :param excel_workbook: ExcelWorkbook实例
        :param sheet: sheet表格内容
        :param sheet_no: sheet索引
        :param start_row_num: row索引
        :param row_del_class: 解析处理row的类，默认是ImportRow
        '''
        super().__init__(excel_workbook, row_del_class, row_validate_func, sheet_no)
        self.__sheet = sheet
        self.start_row_num = start_row_num
        self.total_row_num = sheet.nrows  # 总行数
        self.total_col_num = sheet.ncols  # 总列数
        self.title_map = {}
        self.error_message_prefix = error_message_prefix or '第{row_num}行'
        self.reader_data_map = {}  # 使用map是为了存入行index，可以采用异步线程存入
        self.error_message_map = {}  # 使用map是为了存入行index，可以采用异步线程存入
        self.merge_cell_value_map = {}  # 存储合并单元格数据的map

    def get_parse_map(self):
        return self.excel_workbook.parse_map

    def set_title_map(self):
        '''
        获取首行的列名和列index的map映射关系
        :return:
        '''
        total_col_num = self.__sheet.ncols
        for col_num in range(0, total_col_num):
            self.title_map[str(self.__sheet.cell(0, col_num).value)] = col_num

    def add_result(self, row_num, row_result):
        '''
        返回结果是{
            'success': deserialize_success,
            'result': reader_data,
            'message_list': error_message_list
        }
        :param row_num: 行数
        :param row_result: 行解析结果
        :return:
        '''
        if row_result.get('success') is True:
            self.reader_data_map[row_num] = row_result.get('result')
        else:
            self.error_message_map[row_num] = row_result.get('message_list')

    def parse_row(self, row_num):
        '''
        行的解析和验证，有自定义验证方法时调用自定义验证方法
        :param row_num:
        :return:
        '''
        result = self.row_del_class(self, self.__sheet.row(row_num), row_num).get_value()
        if self.row_validate_func is not None:
            error_message_list = self.row_validate_func(row_num, self.__sheet.row(row_num), self.get_parse_map())
            if error_message_list is not None and len(error_message_list) != 0:
                result['success'] = False
                # 拼接错误信息
                for index, message in enumerate(error_message_list):
                    error_message_list[index] = '{}{}'.format(self.error_message_prefix.format(row_num=row_num), message)
                result['message_list'].extend(error_message_list)
        self.add_result(row_num, result)

    def sync_parse(self):
        '''
        同步方式进行解析
        :return:
        '''
        for row_num in range(self.start_row_num, self.excel_workbook.end_row_num or self.total_row_num):
            self.parse_row(row_num)

    def thread_parse(self):
        '''
        异步模式进行解析
        :return:
        '''

        def work_func(that, rn):
            that.parse_row(rn)
        work_list = []
        # 创建线程池
        with ThreadPoolExecutor(max_workers=self.excel_workbook.max_workers) as executor:
            for row_num in range(self.start_row_num, self.excel_workbook.end_row_num or self.total_row_num):
                future = executor.submit(work_func, self, row_num)
                work_list.append(future)
            # 等待完成
            wait(work_list)

    def del_merged_cells(self):
        '''
        处理合并单元格
        '''
        for (start_row, end_row, start_col, end_col) in self.__sheet.merged_cells:
            for row in range(start_row, end_row):  # 起始行和结束行为前闭后开关系，从0开始算
                for col in range(start_col, end_col):  # 起始列和结束列为前闭后开关系，也是从0开始算
                    if row != start_row or col != start_col:
                        self.merge_cell_value_map[(row, col)] = self.__sheet.cell(start_row, start_col)

    def parse_import(self):
        '''
        开始解析的函数，相当于解析的入口
        :return:
        '''
        self.del_merged_cells()
        self.set_title_map()
        if self.excel_workbook.sync is True:
            self.sync_parse()
        else:
            self.thread_parse()

    def get_value(self):
        return Result(
            result=sort_dict_data(self.reader_data_map),
            error_message_list=sort_dict_list_data(self.error_message_map)
        )
