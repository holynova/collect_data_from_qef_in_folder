# -*- coding: utf-8 -*-
from openpyxl import Workbook,load_workbook

import os,datetime,xlrd
current_dir = os.path.dirname(os.path.abspath(__file__))
# input_dir = current_dir+"\\output V3-convert to xls"
input_dir = r'C:\Users\adam\Desktop\xls'
arr_pack = []
for root, dirs, files in os.walk(input_dir, topdown=False):
    for name in files:
        # print( 'files----'+os.path.join(root, name))
        arr_pack.append(os.path.join(root, name))

def find_cell_value(sht,cell_str):
	# print "%s=(%s,%s)" %(cell_str,find_row(cell_str),find_col(cell_str)),
	return sht.cell_value(rowx = find_row(cell_str),colx=find_col(cell_str))

def find_col(cell_str):
	col_str = cell_str.lower()[:1]
	return 'abcdefghijklmnopqrstuvwxyz'.find(col_str)

def find_row(cell_str):
	row_str = cell_str.lower()[1:]
	return int(row_str)-1


print 'name,gm1_p,gm2_p,gm3_p,cm1_p'
str_show = ''
i=1
for pack_name in arr_pack:
	book = xlrd.open_workbook(pack_name)
	sht = book.sheet_by_index(0)

	gm1_p = find_cell_value(sht,'E17')
	gm2_p = find_cell_value(sht,'E22')
	gm3_p = find_cell_value(sht,'E28')
	cm1_p = find_cell_value(sht,'E30')
	str_show += "%s,%.4f,%.4f,%.4f,%.4f\n" %(pack_name,gm1_p,gm2_p,gm3_p,cm1_p)
	# str_show += str(i)+":"+pack_name+'\n' +str(find_cell_value(sht,'b1')) +'\n'
	i += 1
print str_show

		

