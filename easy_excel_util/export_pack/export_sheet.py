# !user/bin/env python3
# -*- coding: utf-8 -*-
# Author: ChaoqiYin
from concurrent.futures import ThreadPoolExecutor, wait


class ExportSheet(object):

    def __init__(self, excel_workbook, workbook, sheet_no, sheet_map, max_workers, row_del_class):
        '''
        init
        :param excel_workbook: BaseWorkbook子类实例
        :param workbook: workbook待写入文件
        :param sheet_no: 表格索引
        :param sheet_map: sheet_map实例
        :param max_workers: max_workers
        :param row_del_class: 解析处理row的类
        '''
        self.excel_workbook = excel_workbook
        self.workbook = workbook
        self.work_sheet = workbook.add_sheet(sheet_map.sheet_name, cell_overwrite_ok=True)
        self.sheet_no = sheet_no
        self.sheet_map = sheet_map
        self.max_workers = max_workers
        self.row_del_class = row_del_class
        self.col_data_map = {}  # {1: {1: [xx, style], 2: [yy, style]}}...

    @property
    def parse_map(self):
        '''
        获取解析字段的map映射
        :return:
        '''
        return self.sheet_map.parse_map

    def add_col_data_to_map(self, col_index, row_number, data, style):
        '''
        将需要合并的数据加入到map中
        :param col_index:
        :param row_number:
        :param data:
        :return:
        '''
        if self.col_data_map.__contains__(col_index) is False:
            self.col_data_map[col_index] = {}
        self.col_data_map[col_index][row_number] = [data, style]

    def merge_same_col(self):
        # 首先对每列的数据进行row_num排序
        for col_index, row in self.col_data_map.items():
            row_nums = list(row.keys())
            row_nums.sort()

            merge_dict = dict(
                start_row=0, end_row=0, start_col=col_index, end_col=col_index
            ) # 合并的start行/end行/start列/end列, style

            last_row_num = 0
            for row_num in row_nums:
                if last_row_num + 1 != row_num or row.get(last_row_num, [-1, None])[0] != row.get(row_num, [-2, None])[0]:
                    # 重置merge_dict
                    merge_dict = dict(
                        start_row=row_num, end_row=row_num, start_col=col_index, end_col=col_index
                    )
                # 判断条件：last_row_num不为None且是相邻两行且两行数据相同
                elif last_row_num + 1 == row_num and row.get(last_row_num, [-1, None])[0] == row.get(row_num, [-2, None])[0]:
                    merge_dict['end_row'] = row_num

                if merge_dict['start_row'] != merge_dict['end_row']:
                    # 到达合并触发条件
                    self.work_sheet.merge(
                        merge_dict['start_row'], merge_dict['end_row'], merge_dict['start_col'], merge_dict['end_col'],
                        row[merge_dict['start_row']][1]
                    )
                # 结束后赋值给last_row_num
                last_row_num = row_num

    def sync_parse(self):
        '''
        同步方式进行解析
        :return:
        '''
        row_num = self.add_title(0)
        for row_data in self.sheet_map.list_data:
            row_num += 1
            self.add_row(row_num, row_data)

    def thread_parse(self):
        '''
        异步模式进行解析
        :return:
        '''
        def work_func(that, rn, data):
            that.add_row(rn, data)

        work_list = []
        # 创建线程池
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            row_num = self.add_title(0)
            for row_data in self.sheet_map.list_data:
                row_num += 1
                future = executor.submit(work_func, self, row_num, row_data)
                work_list.append(future)
            # 等待完成
            wait(work_list)

    def add_title(self, row_num):
        '''
        添加sheet标题栏，后期可能会添加自定义抬头，所以传入row_num, 有col_name的情况才写入title
        :return:
        '''
        has_title = False
        for export_field_name, export_field in self.parse_map.items():
            if export_field.col_name is not None:
                has_title = True
                break
        if has_title is True:
            self.row_del_class(self, row_num, []).write_title()
        else:
            row_num -= 1
        return row_num

    def add_row(self, row_num, row_data):
        self.row_del_class(self, row_num, row_data).write_row()

    def parse_export(self):
        if self.max_workers is None:
            self.sync_parse()
        else:
            self.thread_parse()
        # 合并相同单元格
        self.merge_same_col()
