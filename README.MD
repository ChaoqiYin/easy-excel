# easy-excel  

**excel操作简化，包含导入和导出功能。省去平时频繁的重复操作，同时支持xls和xlsx**  

安装方式：  
1. 下载文件内的easy_excel_util包使用
2. pip install easy-excel-util

+ 最简单的使用方法
    ```python
    # 数据导出excel
    from easy_excel_util import Builder, ExportField
  
    data = [{'id': 1, 'name': '姓名1', 'age': 18}, {'id': 2, 'name': '姓名2', 'age': 20}]
    
    Builder.build_export().sheet(
        0, data, dict(
            id=ExportField(index=0, col_name='id'), name=ExportField(1, col_name='姓名'), age=ExportField(2, col_name='年龄')
        )
    ).do_export('/User/ChaoqiYin/excel1.xls')
  
    # 导入excel读取数据
    from easy_excel_util import Builder, ImportField

    r = Builder.build_import('/User/ChaoqiYin/excel1.xls').do_import(dict(
        id=ImportField(0), name=ImportField(1), age=ImportField(2)
    ))
    for field in r.result:
        print(field.id, field.name, field.age)
    ```

+ 高级设置  
    首先需要引入Builder类，该类能进行一些导入或导出需要使用的全局统一配置
    ```python
    from easy_excel_util import Builder
    ```
    Builder提供了以下几种方法进行全局的配置：  
    * add_export_style(style, title_style=None)：添加一个全局的导出样式，根据导出excel类型不同使用不同的style    
        ```python
        # xlsx类型的excel样式, 样式必须是一个dict，key是openpyxl库中cell的样式对应的属性名，value是对应的属性值
        from openpyxl.styles import Font
        
        Builder.add_export_style({
            'font': Font(name='Arial', size=14),
        }, title_style={
            'font': Font(name='Arial', size=18, bold=True),
        })
      
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
        
        Builder.add_export_style(DEFAULT_STYLE)
        ```  
    * add_export_converter(data_type_class, func)：添加一个全局的导出时数据转换方法，data_type_class是需要待转换的python数据类型，根据data_type_class的不同会对对应的python数据进行转换 
        ```python
        builder.add_export_converter(type(float), func=lambda float_value: str(float_value))  # func接收一个value参数，该值为待导出到excel的数据
        ```  
&nbsp;
+ excel导入  
    导入时需要构建一个导入的实例，设置导入解析参数，然后执行导入动作
    ```python
    from easy_excel_util import ImportField
  
    # 直接指定文件路径导入
    import_workbook = builder.build_import(file_path='filepath')  # file_path指定需要导入的excel文件的路径，只能是xls文件
    # 或者指定文件内容导入
    with open('filepath') as f:
        import_workbook = builder.build_import(file_path=None, file_content=f.read())  # file_content指定需要导入的excel的文件的内容，设置了file_content会导致file_path会无效
    # 设置解析对应参数
    import_workbook.set_parse_map(
        username=ImportField(index=0),  # index为excel中待解析的列索引，username即关键字参数的key值，即为转换后的行属性名，可自己任意定义
        work_number=ImportField(index=1)
    )
    # 执行导入动作，取得导入的结果
    result = import_workbook.do_import()
    print(result.success)  # 导入过程是否全部通过
    print(result.error_message_list)  # 导入过程中的错误信息
    print(result.result)  # 导入后的结果列表
    ```

&nbsp;
+ excel导出  
    导出时需要构建一个导出的实例，设置导出解析参数，然后执行导出动作
    ```python
    export_workbook = builder.build_export(file_path_or_stream='filepath')  # file_path_or_stream指定需要保存的excel文件的路径或stream流，只能导出xls文件，stream流可用于http请求中下载文件使用
    export_workbook.set_sheet_map()
    ```
