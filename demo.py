# -*- coding:utf-8 -*-
import xlrd
import os
from scripts import Action
from UI import MianWnd
import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *




def wanna_say():
	print (u'第一次写，随便写了点功能，后续可能会加入:\n[代理，CS，远程更新，并发调优，数据集中存储，界面交互，日志，反爬虫等模块]\t异常处理也会做的更仔细')
	print (u'WARNING:因公司网络限制无法下载媒体文件!公司外可以随意整。')
    
 
def	simple_ui():
	#config = ConfigDealer()
	wanna_say()
	app = QApplication(sys.argv)
	mian_window = QMainWindow()
	ui = MianWnd.Ui_EasyEngBook()
	ui.setupUi(mian_window)
	mian_window.show()
	sys.exit(app.exec())
#mark
	root_path = "input/"
	file_lst = os.listdir(root_path)
	for filename in file_lst:
		act = Action(root_path+filename)
		act.write_info_into_excel()
	#txtdler = text_dealer(path)
	#exc = excel_dearler(config,txtdler,path)
	#net = net_work(config)




def test():
	book = xlrd.open_workbook("sample.xls", formatting_info=True)
	sheets = book.sheet_names()
	print("sheets are:", sheets)
	for index, sh in enumerate(sheets):
		sheet = book.sheet_by_index(index)
		print("Sheet:", sheet.name)
		rows, cols = sheet.nrows, sheet.ncols
		print("Number of rows: %s Number of cols: %s" % (rows, cols))
		for row in range(rows):
			for col in range(cols):
				print("row, col is:", row + 1, col + 1,)
				thecell = sheet.cell(row, col)  # could get 'dump','value', 'xf_index'
	print(thecell.value,)
	xfx = sheet.cell_xf_index(row, col)
	xf = book.xf_list[xfx]
	bgx = xf.background.pattern_colour_index

if __name__ == '__main__':
	simple_ui()
	#test()
	#net();
	input("Enter enter key to exit...")