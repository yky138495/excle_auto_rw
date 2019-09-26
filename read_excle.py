import requests
import time
import os
import json
import re
import xlrd
from datetime import date,datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

excle_file = '公司名字频统计.xlsx'

def read_excel():
	wb = xlrd.open_workbook(filename = excle_file)
	print(wb.sheet_names())#获取所有表格名字
	sheet1 = wb.sheet_by_index(0)#通过索引获取表格
	# sheet2 = wb.sheet_by_name('年级')#通过名字获取表格
	# print(sheet1,sheet2)

	print(sheet1.name,sheet1.nrows,sheet1.ncols)

	rows = sheet1.row_values(0)#获取行内容

	cols = sheet1.col_values(0)#获取列内容

	print(rows)

	# print(cols)

	# print(sheet1.cell(1,0).value)#获取表格里的内容，三种方式
	# print(sheet1.cell_value(1,0))
	# print(sheet1.row(1)[0].value)


if __name__ == "__main__":
   read_excel()


