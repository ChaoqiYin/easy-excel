# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
import xlwt

from xlrd.xldate import xldate_as_datetime

# 转换类对应
EMPTY = 0
TEXT = 1
NUMBER = 2
DATE = 3
BOOLEAN = 4
ERROR = 5
BLANK = 6

# 设置边框
borders = xlwt.Borders()
# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1

# 单元格默认样式
DEFAULT_STYLE = xlwt.XFStyle()
al = xlwt.Alignment()
al.wrap = 1  # 自动换行
al.horz = 0x02  # 设置水平居中
al.vert = 0x01  # 设置垂直居中
DEFAULT_STYLE.alignment = al
DEFAULT_STYLE.borders = borders

# 单元格默认title样式
DEFAULT_TITLE_STYLE = xlwt.XFStyle()
al2 = xlwt.Alignment()
al2.wrap = 1
al2.horz = 0x02  # 设置水平居中
al2.vert = 0x01  # 设置垂直居中
DEFAULT_TITLE_STYLE.alignment = al2
font = xlwt.Font()  # 为样式创建字体
# 字体大小，11为字号，20为衡量单位
font.height = 14 * 20
font.bold = True  # 黑体
DEFAULT_TITLE_STYLE.font = font  # 设定样式
DEFAULT_TITLE_STYLE.borders = borders


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
