# easy-excel  

**excel操作简化，包含导入和导出功能。支持配置字段转换方法、选择异步或同步解析、自定义行校验方法、自定义行校验类、配置报错信息前缀、str字段根据格式转换为datetime格式**  

安装方式：  
1. 下载文件内的easy_excel_util包使用
2. pip install easy-excel-util

用法：  
+ 通用配置  
    首先需要实例化一个Builder类来进行下一步导入或导出操作，可以设置为一个全局实例，该实例能进行一些导入或导出需要使用的统一配置
    ```python
    from easy_excel_util import Builder
    
    build = Builder()
    ```
    Builder的实例提供了以下几种方法进行全局的配置：  
    * add_import_converter(converter_key, func)：添加一个全局的导入时数据转换方法，根据converter_key的不同会对对应的excel表格内数据进行转换    
        ```python
        from easy_excel_util import EMPTY, TEXT, NUMBER, DATE, BOOLEAN, ERROR, BLANK
        
        builder.add_import_converter(NUMBER, func=lambda value: str(value))  # func接收一个value参数，该值为excel单元格内的数据
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
