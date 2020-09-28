# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import datetime


class ImportCell(object):

    def __init__(self, import_row, index, import_field):
        '''
        init
        :param import_row: ImportRow实例
        :param index: 单元格列位置，从0开始
        :param import_field: 导入字段设置
        '''
        import_sheet = import_row.import_sheet
        self.excel = import_sheet.excel
        self.import_row = import_row
        self.cell = import_sheet.merge_cell_value_map.get((import_row.row_num, index), self.excel.cell(import_row.row_num, index))  # 判断是否属于合并单元格
        self.index = index
        self.import_field = import_field


    def get_error_message_template(self):
        '''
        获取错误消息模板
        :return:
        '''
        return '{}{}{}'.format(self.import_row.import_sheet.error_message_prefix, '{col_name}', '{message}')

    def get_value(self):
        '''
        获取value
        :return:
        '''
        return self.deserialize_cell()

    def get_converter(self, ctype):
        '''
        获取当前应该使用的转换方法
        :param ctype: 当前cell的数据类型
        :return:
        '''
        # 如果字段本身设置了转换方法，则优先使用此转换方法
        if self.import_field.converter is not None:
            return self.import_field.converter
        # 匹配excel_workbook的转换方法，其中excel_workbook中设置的会覆盖builder中设置的转换方法
        return self.import_row.import_sheet.excel_workbook.converters.get(ctype, None)

    def converter_del_value(self, value, ctype):
        '''
        根据转换方法处理数据
        :param value: 当前cell的数据
        :param ctype: 当前cell的数据类型
        :return:
        '''
        converter = self.get_converter(ctype)
        if converter is None:
            return value
        else:
            return converter(value)

    def check_required_cell(self, value):
        '''
        进行非空必填验证，返回错误信息
        :param value:
        :return:
        '''
        message = None
        if self.import_field.required_message is not None and value is None:
            message = self.get_error_message_template().format(
                row_num=self.import_row.row_num + 1,  # 索引数加1
                col_name=self.import_field.name,
                message=self.import_field.required_message
            )
        return message

    def format_str_to_datetime(self, value):
        '''
        根据字段format转换str为datetime，转换失败会报错
        :param value:
        :return:
        '''
        if self.import_field.datetime_format is None:
            return value
        return datetime.datetime.strptime(str(value), self.import_field.datetime_format)

    def deserialize_cell(self):
        '''
        处理cell的value数据，先处理excel格式为python格式，然后根据是否有转换方法，执行转换方法
        ctype类型关系
            0: 'empty',
            1: 'text',
            2: 'number',
            3: 'xldate',
            4: 'bool',
            5: 'error',
            6: 'blank'
        :return:
        '''
        if self.cell.ctype == 3:  # 日期格式数据，需要转换为datetime
            value = self.excel.del_datetime(self.cell.value)
        elif self.cell.ctype == 4:  # 布尔类型，传入时是0或1，转换为True或False
            value = self.cell.value == 1
        elif self.cell.ctype in (0, 5, 6):  # 空/错误类型/空白，直接设置为None
            value = None
        else:
            value = self.cell.value
        message = self.check_required_cell(value)
        if message is None:
            value = self.converter_del_value(value, self.cell.ctype)
            try:
                value = self.format_str_to_datetime(value)
            except Exception as e:
                # 不能转换格式的直接忽略不转换
                pass
        # 格式转换处理
        return {
            'success': message is None,
            'result': value,
            'message': message
        }
