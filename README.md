��ʶ�밲װ
����ο��ٷ�ʹ���ĵ���
http://openpyxl.readthedocs.io/en/default/usage.html


Openpyxl is a Python library for reading and writing Excel 2010 xlsx/xlsm/xltx/xltm files.

$pip install openpyxl


һ���򵥴�������

from openpyxl import Workbook 

wb = Workbook()

# ���� worksheet

ws = wb.active

# ���ݿ���ֱ�ӷ��䵽��Ԫ����

ws['A1'] = 32

# ���Ը����У��ӵ�һ�п�ʼ����

ws.append([1, 2, 3])

# Python ���ͻᱻ�Զ�ת��

import datetime

ws['A3'] = datetime.datetime.now().strftime("%Y-%m-%d")

# �����ļ�

wb.save("sample.xlsx") 

