# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
# 转换类对应
EMPTY = 0
TEXT = 1
NUMBER = 2
DATE = 3
BOOLEAN = 4
ERROR = 5
BLANK = 6


def sort_dict_data(dict_data, reverse=False):
    '''
    对字典数据进行key的排序处理
    :param dict_data: 字典数据
    :param reverse: 是否降序
    :return: map的values的列表
    '''
    if len(dict_data) == 0:
        return []
    # 进行排序处理
    keys = list(dict_data.keys())
    keys.sort(reverse=reverse)
    return [dict_data.get(key) for key in keys]


def sort_dict_list_data(dict_list_data, reverse=False):
    '''
    对字典-list格式数据进行key的排序处理，并将list整合为一个list，例: {1: ['第1行xx有错', '第1行yy有错'], 2: ['第2行xx有错', '第2行yy有错']}
    :param dict_list_data: 字典数据
    :param reverse: 是否降序
    :return: map的values的列表
    '''
    if len(dict_list_data) == 0:
        return []
    keys = list(dict_list_data.keys())
    keys.sort(reverse=reverse)
    init_list = []
    for key in keys:
        init_list.extend(dict_list_data.get(key))
    return init_list
