# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


class ReaderData(object):
    '''
    excel导出后数据类
    '''
    def __init__(self, row_index):
        '''
        初始化只给与行的index信息，剩余的字段值全部通过setattr()赋予
        :param row_index:
        '''
        self.row_index = row_index
