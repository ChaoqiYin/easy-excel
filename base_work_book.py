# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


class BaseWorkBook(object):
    def __init__(self, file_path, converters):
        self.file_path = file_path  # 文件路径
        self.converters = converters  # 转换类
        self.sync = True  # 是否为同步模式
        self.max_workers = 3  # 异步最大线程数

    def thread(self, max_workers):
        '''
        开启多线程模式
        :param max_workers 最大线程数
        :return:
        '''
        self.sync = True
        self.max_workers = max_workers
        return self

    def _check_map_index_unique(self, parse_map):
        '''
        验证index是否唯一
        :param parse_map:
        :return:
        '''
        index_list = []
        for field_name, build_field in parse_map.items():
            # 设置field的自身__name属性
            build_field.name = field_name
            if build_field.index in index_list:
                raise Exception("the index {index} of field '{name}' is not unique".format(index=build_field.index, name=field_name))
            index_list.append(build_field.index)