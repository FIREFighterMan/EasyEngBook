# -*- coding:utf-8 -*-
import time
import io
import re
import xlrd
import xlwt
from xlutils.copy import copy
import xlutils
import requests #导入需要的包
import json
import sys,os
from datetime import date,datetime

import urllib3 as urllib2
import lxml.html as HTML


class System_Thing:
    def __init__(self):
        self.path = self.get_path()
        
    def get_path(self):
        tail = str(r'/')
        return sys.path[0]+tail

    def file_is_exist(self,filename):
        return os.path.exists(self.path+filename)

def get_word_origin(word):
	url = 'https://www.youdict.com/w/'
	url += word
	headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
	#支持断网续传
	try:
		req = urllib2.Request(url, headers=headers)
		content = urllib2.urlopen(req).read()
		if isinstance(content, unicode):
			pass
		else:
			content = content.decode('utf-8')
		htmlSource = HTML.fromstring(content)
		#US_phonetic_symbol  =  htmlSource.xpath(r"//span[@class='en-US']/text()")
		#详细中文翻译
		translation=htmlSource.xpath('"//*[@id="yd-word-meaning"]/ul/li"')
		print translation
		#美国音标
		phonetic_symbol =   htmlSource.xpath('"//*[@id="yd-word-pron"]/span[2]/text()"')
		print phonetic_symbol
		#词源中文翻译
		simple_ch =   htmlSource.xpath('"//*[@id="yd-ciyuan"]/span/text()"')
		print simple_ch
		#中文词源
		origin_ch =   htmlSource.xpath('"//*[@id="yd-ciyuan"]/p"')
		print origin_ch
		#英文词源
		origin_en =   htmlSource.xpath('"//*[@id="yd-etym"]/dl/dd "')
		print origin_en
		#第一个例句
		exst =   htmlSource.xpath('"//*[@id="yd-liju"]/dl[1]/dt"')
		print exst
		CH_word_origin  =  htmlSource.xpath('"//span[@class="ciyuan-title"]/text()"')
		#US_word_origin  = htmlSource.xpath(r"//span[@class='ciyuan-title']/text()")
	except urllib2.URLError,e:
		print e
		time.sleep(4)
		US_phonetic_symbol = get_word_origin(word)
	else:
		pass  #其他异常的处理
	return US_phonetic_symbol
	
def get_wordmean(word):
    url = 'http://www.iciba.com/'
    url += word
    headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
#支持断网续传
    try:
        req = urllib2.Request(url, headers=headers)
        content = urllib2.urlopen(req).read()
        if isinstance(content, unicode):
            pass
        else:
            content = content.decode('utf-8')
        htmlSource = HTML.fromstring(content)
        mean =  htmlSource.xpath(r"//span[@class='prop']/text()|//li[@class='clearfix']/p/span/text()")
        #phonetic_symbol = htmlSource.xpath(r"//span[@class='base-top-voice']/text()")
        #print phonetic_symbol
    except urllib2.URLError,e:
        print e
        time.sleep(4)
        mean = get_wordmean(word)
    else:
        pass  #其他异常的处理
    return mean

def test(path):
	# 打开文件
	workbookr = xlrd.open_workbook(path+'.xls')
	sheetr = workbookr.sheet_by_index(0) # sheet索引从0开始
	wb = copy(workbookr)
	table = wb.get_sheet(0)
	for row in range(2,sheetr.nrows):
		word = sheetr.cell(row,1).value.encode('utf-8')
		#sentence = self.txt_dealer.search_sentence_by_word(word)
		meanning = get_wordmean(word)
		table.write(row,4,meanning)
		row += 1
		wb.save(path+'.xls') #保存文件

def check_contain_num(check_str):
    flag = False
    for ch in check_str.decode('utf-8'):
        if u'9' >= ch and ch >= u'0':
            flag =  True
    return flag
	
class text_dealer:
	def __init__(self, path):
		self.path = path
		self.map_word_frequency = dict()
		#sentence = re.compile(r"\.+")
		with io.open(self.path, encoding="utf-8") as f:
			self.data = f.read()
			self.sentences = [s.lower() for s in re.findall(r"\.+", self.data)]
			self.words = [s.lower() for s in re.findall(r"\w+", self.data)]
		for word in self.words:
			self.map_word_frequency[word] = self.map_word_frequency.get(word, 0) + 1
		
	def search_sentence_by_word(self,word):
		sentence = re.findall(r'[^.]*?{}[^.]*?.'.format(word), self.data)
		for i, r in enumerate(sentence, 1):
			s = r
			return r+'.'
		return 'NULL'
		
	def get_map_word_frequency(self):
		return self.map_word_frequency
		
	def write_in_txt(self):
		self.fopen = open(self.path+".txt",'w')
		for word in self.map_word_frequency :
			self.fopen.write(str(word)+'\n')
		self.fopen.close()
        
        
class excel_dearler:
	def __init__(self, path,data,txt_dealer):
		self.file_name = path+'.xls'
		#self.full_path = System_Thing().get_path()+path+'.xls'
		self.full_path = self.file_name
		self.data = data
		self.title_style = self.set_title_style('Times New Roman',220,True)
		self.write_word_frequency()
		self.write_word_meaning()
		self.txt_dealer = txt_dealer
		
	def save(self,workbook):
		try:
			workbook.save(self.file_name)
		except:
			print('please close the excle')
			time.sleep(3)
			self.save(workbook)
		
	def write_word_meaning(self):
		workbookr = xlrd.open_workbook(self.file_name)
		sheetr = workbookr.sheet_by_index(0) # sheet索引从0开始
		
		print (r'Start to write in meanning........\n')
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		for row in range(2,sheetr.nrows):
			# yisi = data = sheetr.cell(row,4).value
			#二次传输时自动识别弥补，不覆盖已有
            #单元格内容不为空
			# if data!='':
				# row = row + 1
				# continue
			word = sheetr.cell(row,1).value.encode('utf-8')
			#ciyuan
			time.sleep(3)
			get_word_origin(word)
			# meanning = get_wordmean(word)
			# if len(meanning) == 0:
				# meanning = 'NULL'
			# sheet.write(row,4,meanning)
			# row+=1
			# workbook.save(self.full_path) #保存文件
		# print ('All the meanning be written\n')
         
	# def read_excel(self):
		# # 打开文件
		# workbookr = xlrd.open_workbook(self.full_path)
		
		# sheetr = workbookr.sheet_by_index(0) # sheet索引从0开始
		# wb = copy(workbookr)
		# table = wb.get_sheet(0)
		# for row in range(2,sheetr.nrows):
			# word = sheetr.cell(row,1).value.encode('utf-8')
			# wm = sheetr.cell(row,4).value.encode('utf-8')
			# #二次传输时自动识别弥补，不覆盖已有
			# if wm：
                # row++
                # continue
			# time.sleep(4)
			# meanning = get_wordmean(word)
			# table.write(row,4,meanning)
			# row += 1
			# wb.save(self.full_path) #保存文件
			
	def set_title_style(self,name,height,bold=False):
		style = xlwt.XFStyle()  # 初始化样式

		font = xlwt.Font()  # 为样式创建字体
		font.name = name # 'Times New Roman'
		font.bold = bold
		font.color_index = 4
		font.height = height

		borders= xlwt.Borders()
		borders.left= 6
		borders.right= 6
		borders.top= 6
		borders.bottom= 6

		style.font = font
		style.borders = borders
		print(r'set_title_style\n')
		return style

	def create_excel(self,sheet_name):
		workbook = xlwt.Workbook() #创建工作簿
		workbook.add_sheet(u'频率统计',cell_overwrite_ok=True) #创建sheet
		workbook.save(self.full_path) #保存文件
		print ('create_excel\n')
        
	def open_excel(self):
		workbook = xlrd.open_workbook(self.full_path)
		return copy(workbook)
        
	#写入原始数据，如表格形式等
	def write_original_data(self):
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		row0 = [u'单词',u'出现频率',u'词根',u'中文翻译']
			#生成第一行
		for i in range(1,len(row0)+1):
			sheet.write(1,i,row0[i-1],self.title_style)
		workbook.save(self.full_path)
		print 'write_original_data\n'
		
	#写excel
	def write_word_frequency (self):
		syst = System_Thing()
		if syst.file_is_exist(self.full_path):
			print('{} is already exists'.format(self.full_path))
			pass
		else:
			self.create_excel('sheet1')
			self.write_original_data()
			workbook = self.open_excel()
			sheet = workbook.get_sheet(0)
		#填写数据word_frequency
			row = 2
			for (key,value) in self.data.items():
				if str.isdigit(key.encode('gbk')) or len(key)<2 or check_contain_num(key) :
					pass
				else:
					sheet.write(row,1,key)
					sheet.write(row,2,value)
					row += 1
				workbook.save(self.full_path) #保存文件
			print ('write_word_frequency\n')
		
if __name__ == '__main__':

	path = "Steve Jobs.txt"

	txtdler = text_dealer(path)
	word_frequency = txtdler.get_map_word_frequency()
	exc = excel_dearler(path,word_frequency,txtdler)

	# test(path)