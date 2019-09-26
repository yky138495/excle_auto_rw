import requests
import time
import os
import json
import re
import xlrd
from datetime import date,datetime
import xlwt
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

excle_file = '公司名字频统计.xlsx'

def create_excle():
	# 创建一个workbook 设置编码
	workbook = xlwt.Workbook(encoding = 'utf-8')
	# 创建一个worksheet
	worksheet = workbook.add_sheet('My Worksheet')

	# 写入excel
	# 参数对应 行, 列, 值
	worksheet.write(1,0, label = 'this is test')

	# 保存
	workbook.save('Excel_test.xls')

def write_test():
	workbook = xlwt.Workbook(encoding = 'ascii')
	worksheet = workbook.add_sheet('My Worksheet')
	style = xlwt.XFStyle() # 初始化样式
	font = xlwt.Font() # 为样式创建字体
	font.name = 'Times New Roman' 
	font.bold = True # 黑体
	font.underline = True # 下划线
	font.italic = True # 斜体字
	style.font = font # 设定样式
	worksheet.write(0, 0, 'Unformatted value') # 不带样式的写入
	worksheet.write(1, 0, 'Formatted value', style) # 带样式的写入
	workbook.save('formatting.xls') # 保存文件
#设置单元格宽度:
def set_cell_height():
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	worksheet.write(0, 0,'My Cell Contents')

	# 设置单元格宽度
	worksheet.col(0).width = 3333
	workbook.save('cell_width.xls')

	#输入一个日期到单元格:
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	style = xlwt.XFStyle()
	style.num_format_str = 'M/D/YY' # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
	worksheet.write(0, 0, datetime.datetime.now(), style)
	workbook.save('Excel_Workbook.xls')


	#向单元格添加一个公式:
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	worksheet.write(0, 0, 5) # Outputs 5
	worksheet.write(0, 1, 2) # Outputs 2
	worksheet.write(1, 0, xlwt.Formula('A1*B1')) # Should output "10" (A1[5] * A2[2])
	worksheet.write(1, 1, xlwt.Formula('SUM(A1,B1)')) # Should output "7" (A1[5] + A2[2])
	workbook.save('Excel_Workbook.xls')


	#向单元格添加一个超链接:
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	worksheet.write(0, 0, xlwt.Formula('HYPERLINK("http://www.google.com";"Google")')) # Outputs the text "Google" linking to http://www.google.com
	workbook.save('Excel_Workbook.xls')


	#合并列和行:
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	worksheet.write_merge(0, 0, 0, 3, 'First Merge') # Merges row 0's columns 0 through 3.
	font = xlwt.Font() # Create Font
	font.bold = True # Set font to Bold
	style = xlwt.XFStyle() # Create Style
	style.font = font # Add Bold Font to Style
	worksheet.write_merge(1, 2, 0, 3, 'Second Merge', style) # Merges row 1 through 2's columns 0 through 3.
	workbook.save('Excel_Workbook.xls')


	#设置单元格内容的对其方式:
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	alignment = xlwt.Alignment() # Create Alignment
	alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
	alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
	style = xlwt.XFStyle() # Create Style
	style.alignment = alignment # Add Alignment to Style
	worksheet.write(0, 0, 'Cell Contents', style)
	workbook.save('Excel_Workbook.xls')


	#为单元格议添加边框:
	# Please note: While I was able to find these constants within the source code, on my system (using LibreOffice,) I was only presented with a solid line, varying from thin to thick; no dotted or dashed lines.
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	borders = xlwt.Borders() # Create Borders
	borders.left = xlwt.Borders.DASHED #DASHED虚线  NO_LINE没有 THIN实线
	borders.right = xlwt.Borders.DASHED
	borders.top = xlwt.Borders.DASHED
	borders.bottom = xlwt.Borders.DASHED
	borders.left_colour = 0x40
	borders.right_colour = 0x40
	borders.top_colour = 0x40
	borders.bottom_colour = 0x40
	style = xlwt.XFStyle() # Create Style
	style.borders = borders # Add Borders to Style
	worksheet.write(0, 0, 'Cell Contents', style)
	workbook.save('Excel_Workbook.xls')


	#为单元格设置背景色:
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('My Sheet')
	pattern = xlwt.Pattern() # Create the Pattern
	pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
	pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
	style = xlwt.XFStyle() # Create the Pattern
	style.pattern = pattern # Add Pattern to Style
	worksheet.write(0, 0, 'Cell Contents', style)
	workbook.save('Excel_Workbook.xls')

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
   create_excle()
   write_test()


