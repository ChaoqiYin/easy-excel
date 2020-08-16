# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


class BaseField(object):
    def __init__(self, index, datetime_format=None, col_name=None, converter=None):
        '''
        init
        :param index: 解析列对应位置，从0开始
        :param datetime_format: 字符串/datetime互相转换的模板，导入时不满足会计入错误
        :param col_name: 列名, 用于报错提示
        :param converter: 字段转换方法
        '''
        self.index = index
        self.datetime_format = datetime_format
        self.col_name = col_name
        self.converter = converter
        self.__name = None  # 列解析后的属性名

    @property
    def name(self):
        return self.col_name or self.__name

    @name.setter
    def name(self, name):
        self.__name = name