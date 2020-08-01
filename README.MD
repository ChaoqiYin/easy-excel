# easy-excel
excel操作简化，包含导入和导出功能
#### 依赖xlrd、xlwt, futures库

## 可使用pip install easy-excel-util

```python
from flask import Flask, request
from excel_util import Builder, ImportField, TEXT

app = Flask(__name__)

build = Builder()


@app.route('/upload', methods=['POST', 'GET'])
def upload():
    if request.method == 'POST':
        f = request.files['file']

        result_map = build.build_import(file_path=None, file_content=f.read()).sheet_info(sheet_no=0, start_row_num=1).set_parse_map(
                    number=ImportField(index=0),
        ).do_import()

        print(result_map.success)
        print(result_map.error_message_list)
        print(result_map.result)
    return "abc"


if __name__ == '__main__':
    app.run()
```
参数作用都在方法里有注释, 支持配置字段转换方法、选择异步或同步解析、自定义行校验方法、自定义行校验类、配置报错信息前缀、str字段根据格式转换为datetime格式
