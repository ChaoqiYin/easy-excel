# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from .import_cell import ImportCell
from .reader_data import ReaderData


class ImportRow(object):
    def __init__(self, import_sheet, row, row_num):
        '''
        init
        :param import_sheet: ImportSheet实例
        :param row: 解析的excel行内容
        :param row_num: 当前行数，从0开始
        '''
        self.import_sheet = import_sheet
        self.row = row
        self.row_num = row_num

    @property
    def total_col_num(self):
        '''
        取得sheet总列数
        :return:
        '''
        return self.import_sheet.total_col_num

    @property
    def parse_map(self):
        '''
        获取解析字段的map映射
        :return:
        '''
        return self.import_sheet.parse_map

    def matching_index_value(self, build_field):
        '''
        根据索引去解析对应的cell
        :param build_field:
        :return:
        '''
        index = build_field.index
        # 索引值在总列数内的情况
        if (index + 1) <= self.total_col_num:
            value = ImportCell(self, self.row[index], index, build_field).get_value()
        else:
            value = None
        return value

    def get_value(self):
        '''
        获取解析后的对象, 反序列化行数据
        :return:
        '''
        error_message_list = []
        deserialize_success = True
        if self.parse_map is None:
            raise Exception("must call the function 'set_parse_map' before read excel!")
        # 遍历找寻对应的字段信息组装为ReaderData对象
        reader_data = ReaderData(self.row_num)
        for field_name, build_field in self.parse_map.items():
            build_field.name = field_name
            result_map = self.matching_index_value(build_field)
            setattr(reader_data, field_name, result_map.get('result'))
            # 根据校验是否成功添加错误信息
            if result_map.get('success') is False:
                deserialize_success = False
                error_message_list.append(result_map.get('message'))
        return {
            'success': deserialize_success,
            'result': reader_data,
            'message_list': error_message_list
        }
