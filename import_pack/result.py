# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin


class Result(object):
    def __init__(self, result, error_message_list):
        self.success = len(error_message_list) == 0
        self.result = result
        self.error_message_list = error_message_list