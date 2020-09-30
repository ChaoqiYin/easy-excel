# easy-excel  
##### version: 2.0+

**excel操作简化，包含导入和导出功能。省去平时频繁的重复操作，同时支持xls和xlsx**  

安装方式：  
需要python3.6以上
1. 下载文件内的easy_excel_util包使用
2. pip install easy-excel-util

+ 最简单的使用方法
    ```python
    # 数据导出excel
    from easy_excel_util import Builder, ExportField
  
    data = [{'id': 1, 'name': '姓名1', 'age': 18}, {'id': 2, 'name': '姓名2', 'age': 20}]
    
    Builder.build_export().sheet(
        index=0, data=data, parse_map=dict(
            id=ExportField(index=0, col_name='id', width=5), 
            name=ExportField(index=1, col_name='姓名', width=15), 
            age=ExportField(index=2, col_name='年龄', width=15)
        )
    ).do_export('/User/ChaoqiYin/excel1.xls')
  
    # 导入excel读取数据
    from easy_excel_util import Builder, ImportField

    r = Builder.build_import('/User/ChaoqiYin/excel1.xls').do_import(dict(
        id=ImportField(index=0), name=ImportField(index=1), age=ImportField(index=2)
    ))
    for field in r.result:
        print(field.id, field.name, field.age)
    ```

+ 全局设置  
    Builder类能进行一些导入或导出需要使用的全局统一配置，提供了以下几种类方法进行全局的配置：  
    * add_export_style(style, title_style=None, xlsx=False)：添加一个全局的导出样式，根据导出excel类型不同使用不同的style, xls和xlsx的样式可以同时存在    
        ```python
        # xlsx类型的excel样式, 样式必须是一个dict，key是openpyxl库中cell的样式对应的属性名，value是对应的属性值
        from openpyxl.styles import Font
        
        Builder.add_export_style(style={
            'font': Font(name='Arial', size=14),
        }, title_style={
            'font': Font(name='Arial', size=18, bold=True),
        }, xlsx=True)
      
        # xls类型的excel样式，样式是一个xlwt库的XFStyle实例
        import xlwt
      
        DEFAULT_STYLE = xlwt.XFStyle()
        al = xlwt.Alignment()
        al.wrap = 1  # 自动换行
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        DEFAULT_STYLE.alignment = al
        font = xlwt.Font()  # 为样式创建字体
        # 字体大小，14为字号，20为衡量单位
        font.name = 'Arial'
        font.height = 10 * 20
        DEFAULT_STYLE.font = font
        
        Builder.add_export_style(style=DEFAULT_STYLE)
        ```  
    * add_export_converter(data_type_class, func)：添加一个全局的导出数据转换方法，data_type_class是需要待转换的python数据类型，根据data_type_class自动匹配数据进行转换  
        ```python
        Builder.add_export_converter(type(float), func=lambda float_value: float_value + 1)  # func接收一个value参数，返回转换后的值
        ```  
    * add_import_converter(converter_key, func)：添加一个全局的导入数据转换方法，converter_key是excel中的数据格式对应key，根据数据格式自动匹配数据进行转换  
        ```python
        from easy_excel_util import EMPTY, TEXT, NUMBER, DATE, BOOLEAN, ERROR, BLANK
      
        Builder.add_import_converter(NUMBER, func=lambda number_value: number_value + 1)  # func接收一个value参数，返回转换后的值
        ``` 
&nbsp;
+ 导出的其他设置  
    * 导出xlsx类型的excel:   
        ```python
        # 默认导出xls格式
        Builder.build_export(xlsx=True)...
        ```
    * build_export()的sheet方法支持多种配置参数:  
        ```python
        def sheet(index, data, parse_map, sheet_name=None, height=40, before=None, after=None,
              style=None, title_style=None, row_del_class=None, max_workers=None):
            '''
            sheet的导出设置
            :param index: sheet索引, 从0开始
            :param parse_map: 解析的dict, key是data中要匹配的key值，value是ExportField类的实例
            :param data: 用list装载导出的数据, 数据可以是字典，也可以是一个对象，只要对象有对应的属性
            :param sheet_name: 表名，不传入时sheet的命名方式为sheet{索引}
            :param height: 行高，默认40，单位为磅
            :param before: 导入第一行前的操作，需要一个func(work_sheet, sheet_no, parse_map, row_del_class)
                           参数分别为内部sheet类的实例、sheet索引、parse_map、行处理类，需要返回一个后续操作进行的起始行索引
            :param after: 导入最后一行后的操作，需要一个func(work_sheet, sheet_no, row_num, parse_map, row_del_class)
                          参数分别为内部sheet类的实例、sheet索引、当前行索引、parse_map、行处理类
            :param style: 该sheet单元格样式，不传则使用全局样式, 可以是一个函数，func(row_num, data, value)
                          参数分别为当前行数，源数据，转换后数据，需要返回一个样式
            :param title_style: 该sheet标题样式，不传则使用全局样式，标题是第一行
            :param row_del_class: 行处理的类，建议设为easy_excel_util.export_pack.export_row.ExportRow的子类
            :param max_workers: 异步线程数，不传入时使用同步模式导出
            :return:
            '''
            ...
            return ...
        ```
    * ExportField支持多种属性设置:
        ```python
        ExportField(index, datetime_format=None, col_name=None, width=20, converter=None, style=None, merge_same=False)
      
        :param index: 导出到sheet的列索引, 从0开始
        :param datetime_format: 如果数据是datetime或者time，根据format转换为对应格式的str
        :param col_name: 第一行的title列名，如果所有字段都不传入col_name，则没有title行
        :param width: 列宽，默认20，单位是字符
        :param converter: 该列的数据转换方法，需要一个func，func接收一个value参数，返回转换后的值，会覆盖全局导出转换
        :param style: 列样式，支持func，参数和返回值同sheet方法中的style，会覆盖sheet中设置的style
        :param merge_same: 该列数据相同的单元格是否合并，默认不合并
        ```
    * 支持一个excel导出多组数据到不同的sheet:
        ```python
        data1 = [{'id': 1, 'name': '姓名1', 'age': 18}, {'id': 2, 'name': '姓名2', 'age': 20}]
        data2 = [{'id': 11, 'name': '姓名11', 'age': 18}, {'id': 12, 'name': '姓名12', 'age': 20}]
    
        Builder.build_export().sheet(
            index=0, data=data1, parse_map=dict(
                id=ExportField(index=0, col_name='id', width=5), 
                name=ExportField(index=1, col_name='姓名', width=15), 
                age=ExportField(index=2, col_name='年龄', width=15)
            )
        ).sheet(
            index=1, data=data2, parse_map=dict(
                id=ExportField(index=0, col_name='id', width=5), 
                name=ExportField(index=1, col_name='姓名', width=15), 
                age=ExportField(index=2, col_name='年龄', width=15)
            )
        ).do_export('/User/ChaoqiYin/excel1.xls')
        ```

&nbsp;
+ 导入的其他设置
    * 支持导入文件路径或file对象，同时兼容xls和xlsx:
        ```python
        # 文件路径
        Builder.build_import('/User/ChaoqiYin/excel1.xls')...
        
        # file对象, request中的file对象也可以
        from flask import request
      
        f = request.files['file']
        Builder.build_import(f)...
        ```
    * 针对单个excel添加数据转换方法:
        ```python
        # 同Builder.add_import_converter, 会覆盖全局设置
        Builder.build_import().add_converter(converter_key, func)...
        ```
    * do_import方法支持多种配置参数:
        ```python
        def do_import(self, parse_map, error_message_prefix='第{row_num}行', sheet_no=0, start_row_num=0, end_row_num=None,
                      row_del_class=None, row_validate_func=None, max_workers=None, title_row=None):
            '''
            导入启动
            :param parse_map: 解析的dict, key是解析excel后result内data中的属性值，value是ImportField类的实例
            :param error_message_prefix: 报错提示的前缀文字, 默认是'第{row_num}'
            :param sheet_no: 解析的表格位置索引，从0开始
            :param start_row_num: 从第几行开始解析，默认0
            :param end_row_num: 到第几行结束，默认是到最后一行数据
            :param row_del_class: 默认的行处理类, 建议是easy_excel_util.import_pack.import_row.ImportRow的子类
            :param row_validate_func: 行验证方法，接收4个参数：（行索引，行原始数据，行转换后的数据，parse_map），返回None或一个list，里面是该行的错误消息，会自动拼接上error_message_prefix
            :param max_workers: 异步线程数，不传入时使用同步模式导入
            :param title_row: 字段col_name自动匹配的标题行索引，传入一个行索引，会根据parse_map中field的col_name去匹配该行的单元格，匹配成功时将index覆盖为对应的单元格index，为None时不匹配
            :return: 返回一个结果对象，有success、result、error_message_list三个属性。success表示导入转换过程中是否有不匹配格式等错误，
                      result是导入转换成功的结果列表，error_message_list是转换有错误时的提示信息列表
            '''
            ...
            return ...
        ```
    * ImportField支持多种属性设置:
        ```python
        ImportField(index, datetime_format=None, col_name=None, converter=None, required_message=None)
      
        :param index: 解析列对应位置，从0开始
        :param datetime_format: excel单元格字符串转换成datetime的格式模板，不满足则计入error_message_list，自动拼接上error_message_prefix
        :param col_name: 列名, 用于报错提示或自动匹配标题行对应的列，匹配需要配合do_import的title_row使用
        :param converter: 该列的数据转换方法，需要一个func，func接收一个单元格value参数，返回转换后的值，会覆盖全局导入转换
        :param required_message: 是否验证非空必填，该列单元格为空时的校验报错信息，计入error_message_list，自动拼接上error_message_prefix
        ```