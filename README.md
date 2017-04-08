初识与安装
详情参考官方使用文档：
http://openpyxl.readthedocs.io/en/default/usage.html


Openpyxl is a Python library for reading and writing Excel 2010 xlsx/xlsm/xltx/xltm files.

$pip install openpyxl


一个简单创建例子

from openpyxl import Workbook 

wb = Workbook()

# 激活 worksheet

ws = wb.active

# 数据可以直接分配到单元格中

ws['A1'] = 32

# 可以附加行，从第一列开始附加

ws.append([1, 2, 3])

# Python 类型会被自动转换

import datetime

ws['A3'] = datetime.datetime.now().strftime("%Y-%m-%d")

# 保存文件

wb.save("sample.xlsx") 

