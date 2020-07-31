# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import xlrd
from .import_sheet import ImportSheet
from ..import_pack.import_row import ImportRow
from ..utils import get_converters_key
from ..base_work_book import BaseWorkBook


def turn_file_to_excel(file_path=None, file_contents=None):
    if file_contents is not None:
        workbook = xlrd.open_workbook(file_contents=file_contents)
    else:
        with open(file_path, encoding='utf8') as f:
            content = f.read()
            workbook = xlrd.open_workbook(file_contents=content)
    return workbook


class ImportWorkbook(BaseWorkBook):
    def __init__(self, file_path, file_content, converters):
        super().__init__(file_path, converters)
        self.file_content = file_content  # 文件内容，文件内容不为None时file_path会失效
        self.sheet_no = 0  # sheet位置，从0开始，默认0
        self.parse_map = None  # 解析map对照
        self.start_row_num = 0  # 从第几行开始解析，默认0开始

    def add_converter(self, converter_key, func):
        '''
        设置转换类
        :param converter_key: 转换类型key
        :param func: 转换方法
        :return:
        '''
        if converter_key not in range(0, 7):
            raise Exception("converter_key must in the dict [0, 1, 2, 3, 4, 5, 6] keys")
        self.converters[get_converters_key(converter_key, True)] = func
        return self

    def set_content(self, file_content=None):
        '''
        设置文件解析路径或内容
        :param file_content: 文件内容
        :return:
        '''
        self.file_content = file_content
        return self

    def sheet_info(self, sheet_no=0, start_row_num=0):
        '''
        sheet解析目标
        :param sheet_no: sheet索引 0开始
        :param start_row_num: 开始行数 0开始
        :return:
        '''
        self.sheet_no = sheet_no
        self.start_row_num = start_row_num
        return self

    def set_parse_map(self, **parse_map):
        '''
        设置解析对应map
        :param parse_map: key为解析字段name，value为BaseField子类
        :return:
        '''
        self._check_map_index_unique(parse_map)
        self.parse_map = parse_map
        return self

    def __validate_before_action(self):
        if self.file_path is None and self.file_content is None:
            raise Exception("file_path and file_content can't all be None!")
        if self.parse_map is None:
            raise Exception("parse_map can't be None!")

    def do_import(self, error_message_prefix=None, row_del_class=None, row_validate_func=None):
        '''
        导入启动方法
        :param error_message_prefix: 报错提示的前缀文字, 默认是'第{row_num}'
        :param row_del_class: 默认的行处理类, 需要是ImportRow的子类
        :param row_validate_func: 行验证方法，返回一个list，里面是该行的错误消息，会自动拼接上error_message_prefix
        :return:
        '''
        rel_row_del_class = ImportRow if row_del_class is None else row_del_class
        self.__validate_before_action()
        workbook = turn_file_to_excel(self.file_path, self.file_content)
        sheet = ImportSheet(self, workbook.sheet_by_index(self.sheet_no), sheet_no=self.sheet_no,
                            start_row_num=self.start_row_num, row_del_class=rel_row_del_class,
                            error_message_prefix=error_message_prefix, row_validate_func=row_validate_func)
        sheet.parse_import()
        return sheet.get_value()
