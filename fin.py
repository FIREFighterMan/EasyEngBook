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
import  urllib
import urllib3 as urllib2
from urllib3 import urlretrieve
import lxml.html as HTML
from multiprocessing import Pool

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import ssl
#import OpenSSL
from requests.packages.urllib3.exceptions import InsecureRequestWarning

class System_Thing:
    def __init__(self):
        self.path = self.get_path()
        
    def get_path(self):
        tail = str(r'/')
        return sys.path[0]+tail

    def file_is_exist(self,filename):
        return os.path.exists(self.path+filename)
def urlretrieve_down_load(word,num):
	try:
		path = r'./html/{}.html'.format(word)
		if os.path.exists(path):
			#print('{} is already exists'.format(path))
			pass
		else:
			#url = r'https://www.youdict.com/w/{}'.format(word)
			#time.sleep(1)
			url = r'https://www.python.org/ftp/python/2.7.12/python-2.7.12.amd64.msi'
			urlretrieve(url, r'python-2.7.12.amd64.msi')
	except:
		print u'异常'
		num += 1
		time.sleep(num)
		urlretrieve_down_load(word,num)

def ssl_get_data(word):
	requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
	r = requests.get('https://www.youdict.com/w/', verify=False)
	txt = r
	print txt
	s =1
	# _create_unverified_https_context = ssl._create_unverified_context
	# ssl._create_default_https_context = ssl._create_unverified_context
	# #context = ssl.SSLContext
	# html =  urllib2.urlopen(r"https://www.youdict.com/w/mensa").read()
	# #sock = ssl.wrap_socket(sock, cert_reqs=ssl.CERT_REQUIRED, ca_certs="cacerts.txt")

	stop =1
def get_word_origin(word):
	ssl._create_default_https_context = ssl._create_unverified_context
	url = 'https://www.youdict.com/w/'
	url += word
	headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
	#支持断网续传
	try:

		req = urllib2.Request(url, headers=headers)
		content = urllib2.urlopen(req).read()
		print content
		if isinstance(content, unicode):
			pass
		else:
			content = content.decode('utf-8')
		htmlSource = HTML.fromstring(content)
		US_phonetic_symbol  =  htmlSource.xpath(r"//span[@class='en-US']/text()")
		print US_phonetic_symbol+'thx'
		CH_word_origin  =  htmlSource.xpath(r"//span[@class='ciyuan-title']/text()")
		#US_word_origin  = htmlSource.xpath(r"//span[@class='ciyuan-title']/text()")
	except Exception,e:
		print e
		time.sleep(4)
		US_phonetic_symbol = get_word_origin(word)
	else:
		pass  #其他异常的处理
	return US_phonetic_symbol

def get_data_by_https():
	date_list = []
	return date_list
def get_data_by_selenium():
	# broswer_path = r'C:\Users\Administrator\AppData\Local\360Chrome\Chrome\Application\360chrome.exe'
	# chrome_options = Options()
	# chrome_options.binary_location = broswer_path

	# chrome = webdriver.Chrome(chrome_options= chrome_options ,executable_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\Application\chromedriver.exe')
	chrome = webdriver.Chrome(executable_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\Application\chromedriver.exe')
	chrome.get(r'https://www.youdict.com/w/coaster')
	#页面源码
	html = HTML.fromstring(chrome.page_source)
	#美式发音
	#pronounce = html.xpath()
	#美式音标
	temp_soundmark = html.xpath(r'//*[@id="yd-word-pron"]/span[2]')
	soundmark = temp_soundmark[0].xpath(r'string()')
	#翻译
	temp_translate= html.xpath(r'//*[@id="yd-word-meaning"]')
	translate = temp_translate[0].xpath(r'string(.)')
	#词根
	temp_root =  html.xpath(r'//*[@id="yd-ciyuan"]/p')
	root = temp_root[0].xpath(r'string(.)')
	#词源
	temp_etymology = html.xpath(r'//*[@id="yd-etym"]')
	etymology = temp_etymology[0].xpath(r'string(.)')
	data_list = [soundmark,translate,root,etymology]
	chrome.close()
	return data_list

def get_wordmean(word):
    url = 'https://www.youdict.com/'
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
		#self.write_word_meaning()
		self.txt_dealer = txt_dealer
		#多进程实验区
		p = Pool(4)
		for i in range(5):
			p.apply_async(self.write_word_meaning(), args=(i,))
		print ("等待所有子进程执行完毕...")
		p.close()
		p.join()
		print ("所有子进程执行完成")
		
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
		print u'start to download'
		for row in range(2,sheetr.nrows):
# .			yisi = data = sheetr.cell(row,4).value
			# #二次传输时自动识别弥补，不覆盖已有
            # #单元格内容不为空
			# if data!='':
				# row = row + 1
				# continue
			word = sheetr.cell(row,1).value.encode('utf-8')
			#ciyuan
			# time.sleep(3)
			urlretrieve_down_load(word,1)
			#meanning = get_wordmean(word)
			#if len(meanning) == 0:
				#meanning = 'NULL'
			#sheet.write(row,4,meanning)
			# row+=1
			# workbook.save(self.full_path) #保存文件
		print ('All the meanning be written\n')
         
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
    get_word_origin(r'show')
	#path = "fin.txt"

	#txtdler = text_dealer(path)
	#word_frequency = txtdler.get_map_word_frequency()
	#exc = excel_dearler(path,word_frequency,txtdler)

	# test(path)
	#urlretrieve_down_load('13',1)
