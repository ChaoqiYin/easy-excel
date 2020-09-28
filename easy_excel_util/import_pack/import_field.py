# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from ..base_field import BaseField


class ImportField(BaseField):
    def __init__(self, index, datetime_format=None, col_name=None, converter=None, required_message=None):
        '''
        init
        :param index: 解析列对应位置，从0开始
        :param datetime_format: 字符串转换成datetime的格式模板，不满足会计入错误
        :param col_name: 列名, 用于报错提示
        :param converter: 字段转换方法
        :param required_message: 是否验证非空必填，为空时的校验报错信息
        '''
        super().__init__(index, datetime_format, col_name, converter)
        self.required_message = required_message  # 验证非空时的校验报错信息
