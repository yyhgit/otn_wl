


#import xlrd
#import shutil
#import os
#
#
#move_list = set()
#
#list_dir = r''
#
#src_dir = r''
#tar_dir = r''


'''
物料维护
'''
import os
#系统包，列举文件夹下文件，移除文件

from xlrd import open_workbook
#excel包，打开excel

from xlutils.copy import copy
#excel包，复制
#import xlwt


dir_1 = 'E:/otn_wl/'
#文件夹路径

for file in os.listdir(dir_1):
#列举文件夹下所有excel文件，循环遍历处理

	file_path = dir_1 + file  #每个文件的完整路径

	rb = open_workbook(file_path,formatting_info=True)#打开excel文件
 
	wb = copy(rb) #拷贝excel给wb
	ws = wb.get_sheet(0) #获取拷贝的excel第0个sheet

	n = rb.sheet_by_index(0).nrows #获取第一sheet页的行数

	for i in range(1,n):
		#循环遍历每一行（0行除外）

		wt_str = rb.sheet_by_index(0).cell(i,13).value
		#原始字符串
		wt_str = 'S'+ wt_str.replace('.','')
		#转换后字符串

		ws.write(i, 13, wt_str)
		#写入转换后字符串

	os.remove(file_path) #移除原来excel

	wb.save(file_path) #保存新的excel


