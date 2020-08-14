# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import xlrd

from .import_sheet import ImportSheet
from ..import_pack.import_row import ImportRow
from ..utils import get_converters_key


def turn_file_to_excel(file_path, file):
    if file is not None:
        workbook = xlrd.open_workbook(file_contents=file.read(), formatting_info=True)
        try:
            file.close()
        except Exception as e:
            pass
    else:
        with open(file_path, encoding='utf8') as f:
            content = f.read()
            workbook = xlrd.open_workbook(file_contents=content, formatting_info=True)
    return workbook


def get_file_path_or_content(p_file):
    file_path = None
    file = None
    if hasattr(file, 'read'):
        file = p_file
    else:
        file_path = p_file
    return file, file_path


class ImportWorkbook(object):
    def __init__(self, file, parse_map, converters, sheet_no=0, start_row_num=0, end_row_num=None):
        p_file_tuple = get_file_path_or_content(file)
        self.file_path = p_file_tuple[0]  # 文件路径
        self.converters = converters  # 转换类
        self.sync = True  # 是否为同步模式
        self.max_workers = 3  # 异步最大线程数
        self.file = p_file_tuple[1]  # 文件内容，文件内容不为None时file_path会失效
        self.sheet_no = sheet_no  # sheet位置，从0开始，默认0
        self.parse_map = parse_map  # 解析map对照
        self.start_row_num = start_row_num  # 从第几行开始解析，默认0开始
        self.end_row_num = end_row_num  # 到第几行停止，默认None
        self.__result_value = None

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

    def set_mode(self, max_workers):
        '''
        开启多线程模式
        :param max_workers 最大线程数
        :return:
        '''
        if max_workers is False or max_workers is None:
            self.sync = True
        else:
            self.sync = False
            self.max_workers = max_workers

    def do_import(self, max_workers=False, error_message_prefix=None, row_del_class=None, row_validate_func=None):
        '''
        导入启动方法
        :param max_workers: 异步最大线程数，为None或False时使用同步模式
        :param error_message_prefix: 报错提示的前缀文字, 默认是'第{row_num}'
        :param row_del_class: 默认的行处理类, 需要是ImportRow的子类
        :param row_validate_func: 行验证方法，返回一个list，里面是该行的错误消息，会自动拼接上error_message_prefix
        :return:
        '''
        self.set_mode(max_workers)
        rel_row_del_class = ImportRow if row_del_class is None else row_del_class
        workbook = turn_file_to_excel(self.file_path, self.file)
        sheet = ImportSheet(self, workbook.sheet_by_index(self.sheet_no), sheet_no=self.sheet_no,
                            start_row_num=self.start_row_num, row_del_class=rel_row_del_class,
                            error_message_prefix=error_message_prefix, row_validate_func=row_validate_func)
        self.__result_value = sheet.get_value()

    @property
    def success(self):
        '''
        解析是否成功
        :return:
        '''
        return self.__result_value.success

    @property
    def error_message_list(self):
        '''
        解析错误信息列表
        :return:
        '''
        return self.__result_value.error_message_list

    @property
    def result(self):
        '''
        解析后的结果列表
        :return:
        '''
        return self.__result_value.result
