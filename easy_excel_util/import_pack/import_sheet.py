# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from concurrent.futures import ThreadPoolExecutor, wait

from .result import Result
from ..utils import sort_dict_data, sort_dict_list_data


class ImportSheet(object):
    def __init__(self, excel_workbook, excel, parse_map, error_message_prefix, sheet_no, start_row_num, end_row_num,
                 max_workers,
                 row_del_class, row_validate_func, title_row):
        '''
        init
        :param excel_workbook: excel_workbook实例
        :param excel: factory的excel类实例
        :param parse_map: 解析的字典
        :param error_message_prefix: 报错提示的前缀文字, 默认是'第{row_num}'
        :param sheet_no: 解析的表格索引
        :param start_row_num: 从第几行开始解析
        :param end_row_num: 到第几行结束
        :param max_workers: 异步最大线程数，为None时使用同步模式
        :param row_del_class: 解析处理row的类，默认是ImportRow
        :param row_validate_func: 行验证方法
        :param title_row: 不为None时，对标题进行验证的row索引，会对此索引的行进行field的col_name验证，如果不匹配，则会直接返回
        '''
        self.excel_workbook = excel_workbook
        self.excel = excel
        self.parse_map = parse_map
        self.error_message_prefix = error_message_prefix
        self.sheet_no = sheet_no
        self.start_row_num = start_row_num
        self.total_row_num = end_row_num or excel.nrows  # 总行数
        self.total_col_num = excel.ncols  # 总列数
        self.max_workers = max_workers
        self.row_del_class = row_del_class
        self.row_validate_func = row_validate_func  # 自定义的行处理方法，未自定义时为None
        self.title_row = title_row
        self.reader_data_map = {}  # 使用map是为了存入行index，可以采用异步线程存入
        self.error_message_map = {}  # 使用map是为了存入行index，可以采用异步线程存入
        self.merge_cell_value_map = {}  # 存储合并单元格数据的map

    def validate_title_index(self):
        '''
        获取指定行的列名和列index的map映射关系
        :return:
        '''
        if self.title_row is not None:
            title_map = {}  # 存储title的名称和index映射关系
            for col_num in range(0, self.total_col_num):
                title_map[str(self.excel.cell(self.title_row, col_num).value)] = col_num

            for export_field_name, export_field in self.parse_map.items():
                # col_name存在的情况，将export_field的index重置为正确的index
                if export_field.col_name is not None and title_map.get(export_field.col_name, None) is not None:
                    setattr(export_field, 'index', title_map.get(export_field.col_name))

    def del_merged_cells(self):
        '''
        处理合并单元格
        '''
        for (start_row, end_row, start_col, end_col) in self.excel.merged_cells:
            for row in range(start_row, end_row):  # 起始行和结束行为前闭后开关系，从0开始算
                for col in range(start_col, end_col):  # 起始列和结束列为前闭后开关系，也是从0开始算
                    if row != start_row or col != start_col:
                        self.merge_cell_value_map[(row, col)] = self.excel.cell(start_row, start_col)

    def parse_row(self, row_num):
        '''
        行的解析和验证，有自定义验证方法时调用自定义验证方法
        结果是{
            'success': deserialize_success,
            'result': reader_data,
            'message_list': error_message_list
        }
        :param row_num:
        :return:
        '''
        result = self.row_del_class(self, row_num).get_value()
        if self.row_validate_func is not None:
            error_message_list = self.row_validate_func(row_num, self.excel.row(row_num), result, self.parse_map)
            if error_message_list is not None and len(error_message_list) != 0:
                result['success'] = False
                # 拼接错误信息
                for index, message in enumerate(error_message_list):
                    error_message_list[index] = '{}{}'.format(self.error_message_prefix.format(row_num=row_num),
                                                              message)
                result['message_list'].extend(error_message_list)
        # 处理结果
        if result.get('success') is True:
            self.reader_data_map[row_num] = result.get('result')
        else:
            self.error_message_map[row_num] = result.get('message_list')

    def sync_parse(self):
        '''
        同步方式进行解析
        :return:
        '''
        for row_num in range(self.start_row_num, self.total_row_num):
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
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            for row_num in range(self.start_row_num, self.total_row_num):
                future = executor.submit(work_func, self, row_num)
                work_list.append(future)
            # 等待完成
            wait(work_list)

    def get_value(self):
        '''
        获取解析结果
        :return:
        '''
        self.del_merged_cells()
        # 匹配title和index
        self.validate_title_index()
        if self.max_workers is None:
            self.sync_parse()
        else:
            self.thread_parse()
        return Result(
            result=sort_dict_data(self.reader_data_map),
            error_message_list=sort_dict_list_data(self.error_message_map)
        )
