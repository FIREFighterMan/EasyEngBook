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

import urllib.request as urllib2
import urllib
import lxml.html as HTML
#import ssl
import configparser as ConfigParser
from threading import Thread
#from gevent import monkey; monkey.patch_all()
#import gevent
#import requests
#from requests.packages.urllib3.exceptions import InsecureRequestWarning
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import ssl

# def net():
# 	requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
# 	r = requests.get('https://www.baidu.com/', verify=False)
# 	print(r.status_code)

def check_contain_num(check_str):
    flag = False
    for ch in check_str.decode('utf-8'):
        if u'9' >= ch and ch >= u'0':
            flag =  True
    return flag
	



def wanna_say():
	print (u'第一次写，随便写了点功能，后续可能会加入:\n[代理，CS，远程更新，并发调优，数据集中存储，界面交互，日志，反爬虫等模块]\t异常处理也会做的更仔细')
	print (u'WARNING:因公司网络限制无法下载媒体文件!公司外可以随意整。')
    
 
def	simple_ui():
	#config = config_dealer()
	wanna_say()
#mark
	root_path = "input/"
	file_lst = os.listdir(root_path)
	for filename in file_lst:
		act = action(root_path+filename)
		act.write_info_into_excel()
	#txtdler = text_dealer(path)
	#exc = excel_dearler(config,txtdler,path)
	#net = net_work(config)	

class action:
	def __init__(self,path):
		try:
			print (u'配置模块初始化...')
			self.config_dealer = config_dealer()
#mark
			self.path = path
			#self.path = self.config_dealer.get_conf(r'book',r'path')
			print (u'文本模块初始化...')
			self.txt_dealer = text_dealer(self.path)
			print (u'网络模块初始化...')
			self.net_work_dealer = net_work(self.config_dealer)
			print (u'Excel模块初始化...')
			self.excel_dealer = excel_dearler(self.config_dealer,self.txt_dealer,self.net_work_dealer,self.path)
		except Exception as error_info:
			print (u'【action_init ERROR】:{}'.format(error_info))
			input("Enter enter key to continue...")
		
	def write_info_into_excel(self):
		try:
			print (u'执行音频下载......')
			#self.excel_dealer.down_load_audio()

			#get = int(raw_input("Insert the audio?('0'for Yes,other for no)"))

			#if get ==0:
			print (u'执行音频插入......')
			#self.excel_dealer.insert_audio()

			# print u'执行信息写入......'
			# #govent = [gevent.spawn(self.excel_dealer.write_word_meaning) for i in range(0,9)]
			# #gevent.joinall(g)
			# ts = [Thread(target=self.excel_dealer.write_word_meaning,args=(i,)) for i in range(0,9)]
			# for t in ts:
				# t.start()
			# for t in ts:
				# t.join()
			self.excel_dealer.write_word_meaning()

			print (u'执行信息优化')
			#self.excel_dealer.optimisation()

			print (u'信息写入完毕！')
		except Exception as error_info:
			print (u'【action_init ERROR】:{}'.format(error_info))
			input("Enter enter key to continue...")

class excel_dearler:
	def __init__(self, config,txt_dealer,net_work,path):
		self.conf = config
		self.txt_dealer = txt_dealer
		self.net_work = net_work

		self.file_name =''
		self.full_path =''
		self.init_config(path)

		self.title_style = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: bold on;')
		self.content_styleA = self.set_title_style(5)
		self.content_styleB = self.set_title_style(4)
		print (u'\t表格样式设置完毕！')
		self.write_word_frequency()
		print (u'\t原始表格数据写入完毕！')
		self.select()
		print (u'根据熟词库将单词筛选完毕！')
		#self.write_all()
		#self.write_word_meaning()

	def select(self):
		#mark 新加入的根据历史记录过滤熟词
		root_path ="output/"
		file_lst = os.listdir(root_path)
		for file in file_lst:
			full_path = root_path+file
			temp = xlrd.open_workbook(self.full_path)
			st = temp.sheet_by_index(0)
			word_work = xlrd.open_workbook(full_path)
			dst_work = copy(temp)
			word_sheet = word_work.sheet_by_index(0)
			dst_sheet = dst_work.get_sheet(0)
			i = word_sheet.nrows
			for row1 in range(0, word_sheet.nrows):
				word = word_sheet.cell(row1, 1).value  # .encode('utf-8')
				if word == '':
					continue
				for row2 in range(0, st.nrows):
					dst = st.cell(row2, 1).value  # .encode('utf-8')
					if word == '':
						continue
					if word == dst:
						empty = r''
						dst_sheet.write(row2, 1, empty)
						# dst_sheet.write(row2, 2, empty)
						break
			self.save(dst_work)
	def insert_youdao_mean_from_loacl(self,word):
		path = './https/youdao/{}.html'.format(word)
		if not os.path.exists(path):
			self.net_work.youdao_down_load(word)
		# 页面源码
		page_source = open(path).read()
		html = HTML.fromstring(page_source)
		# 美式音标
		temp_soundmark = html.xpath(self.conf.get_conf(r'youdao', r'path_soundmark'))
		if len(temp_soundmark) == 0:
			soundmark = r'NULL'
		else:
			soundmark = temp_soundmark[0].xpath(r'string()')
		# 翻译
		temp_translate = html.xpath(self.conf.get_conf(r'youdao', r'path_translate'))
		if len(temp_translate) == 0:
			translate = r'NULL'
		else:
			translate = temp_translate[0].xpath(r'string(.)')

		data_list = [soundmark, translate]
		return data_list
	def insert_mean_from_local(self,word):
		path = './https/{}.html'.format(word)
		if not os.path.exists(path):
			self.net_work.https_down_load(word)
		# 页面源码
		page_source = open(path).read()
		html = HTML.fromstring(page_source)
		# 美式发音
		# pronounce = html.xpath()
		# 美式音标
		temp_soundmark = html.xpath(self.conf.get_conf(r'web', r'path_soundmark'))
		if len(temp_soundmark) == 0:
			soundmark = r'NULL'
		else:
			soundmark = temp_soundmark[0].xpath(r'string()')
		# 翻译
		temp_translate = html.xpath(self.conf.get_conf(r'web', r'path_translate'))
		if len(temp_translate) == 0:
			translate = r'NULL'
		else:
			translate = temp_translate[0].xpath(r'string(.)')
		# 词根
		temp_root = html.xpath(self.conf.get_conf(r'web', r'path_root'))
		if len(temp_root) == 0:
			root = r'NULL'
		else:
			root = temp_root[0].xpath(r'string(.)')
		# 词源
		temp_etymology = html.xpath(self.conf.get_conf(r'web', r'path_etymology'))
		if len(temp_etymology) == 0:
			etymology = r'NULL'
		else:
			etymology = temp_etymology[0].xpath(r'string(.)')

		data_list = [soundmark, translate, root, etymology]
		# web_driver.quit()
		return data_list

	def insert_audio(self):#重复处理后会丢失超链接和背景色
		workbookr = xlrd.open_workbook(self.file_name,formatting_info=True)
		sheetr = workbookr.sheet_by_index(0)  # sheet索引从0开始
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		#区分超链接背景色,重复写入超链接会造成空
		pattern = xlwt.Pattern()
		pattern.pattern = xlwt.Pattern.SOLID_PATTERN
		color = int(self.conf.get_conf('audio', 'color'))
		pattern.pattern_fore_colour = color
		style = xlwt.XFStyle()
		style.pattern = pattern

		counter = 0
		for row in range(2, sheetr.nrows):
			# 二次传输时自动识别弥补，不覆盖已有
			word = sheetr.cell(row, 1).value.encode('utf-8')
			# 若含有超链接则跳过(此处用颜色区别)
			xfx = sheetr.cell_xf_index(row, 1)
			xf = workbookr.xf_list[xfx]
			bgx = xf.background.pattern_colour_index
			if bgx == color:
				row +=1
				continue
			prepath = self.conf.get_conf(r'audio', 'path')
			path = prepath + word+'.mpeg'
			if os.path.exists(path):
				hyper = u'HYPERLINK("{}";"{}")'.format(path, word)
				sheet.write(row, 1, xlwt.Formula(hyper),style)
				counter +=1
			row += 1
		self.save(workbook)
		print (u'插入{}条音频，共{}条数据'.format(counter,sheetr.nrows))

	def init_config(self,path):
		self.file_name = path+r'.xls'
		self.full_path = self.file_name
		
	def save(self,workbook):
		try:
#mark
			root_path = 'output/'
			workbook.save(root_path+self.full_path)
		except:
			print('please close the excle')
			time.sleep(4)
			self.save(workbook)
		
	def write_word_meaning(self):
		workbookr = xlrd.open_workbook(self.file_name)
		sheetr = workbookr.sheet_by_index(0) # sheet索引从0开始
		
		print (u'开始写入网页数据....（该处理时间较长，停一会儿就可以看效果,随时可以关闭程序）')
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		for row in range(2,sheetr.nrows):
			value = sheetr.cell(row,4).value
			#二次传输时自动识别弥补，不覆盖已有
            #单元格内容不为空
			if value!='':
				#row += 10
				continue
			word = sheetr.cell(row,1).value.encode('utf-8')
			if word == '':
				continue
			#time.sleep(3)
			#data_list = self.net_work.get_data_by_selenium(word)
			#data_list = self.net_work.get_data_by_https(word)
			data_list = self.insert_youdao_mean_from_loacl(word)
			#data_list = self.insert_mean_from_local(word)
			sentence = self.txt_dealer.search_sentence_by_word(word)
			for cnt in range(0,2):
				info = data_list[cnt]
				sheet.write(row,cnt+3,info)
			sheet.write(row,7,sentence)
		self.save(workbook)
			#row+=1
			#if row % 10==0:
				#self.save(workbook)
				#print u'已处理{}条数据'.format(row-1)
		print ('All the meanning be written\n')
  
			
	def set_title_style(self,bg_clolor =False):
		# 初始化样式
		style = xlwt.XFStyle()
		#创建字体
		temp_font_color_index = self.conf.get_conf(r'excel', 'style_font_color_index')
		temp_font_name = self.conf.get_conf(r'excel', 'style_font_name')
		temp_font_bold = self.conf.get_conf(r'excel', 'style_font_bold')
		temp_font_height = self.conf.get_conf(r'excel', 'style_font_height')
		font = xlwt.Font()
		font.name = temp_font_name
		font.bold = temp_font_bold
		font.color_index = temp_font_color_index
		font.height = temp_font_height
		#创建下框线
		temp_border_left = self.conf.get_conf(r'excel', 'style_border_left')
		temp_border_right = self.conf.get_conf(r'excel', 'style_border_right')
		temp_border_top = self.conf.get_conf(r'excel', 'style_border_top')
		temp_border_bottom = self.conf.get_conf(r'excel', 'style_border_bottom')
		borders= xlwt.Borders()
		borders.left= temp_border_left
		borders.right= temp_border_right
		borders.top= temp_border_top
		borders.bottom= temp_border_bottom

		#创建背景颜色
		if bg_clolor != False:
			temp_bg_color_index = bg_clolor
		else:
			temp_bg_color_index = self.conf.get_conf(r'excel','style_bg_color_index')
		#temp_bg_pattern = self.conf.get_conf(r'excel', 'style_bg_pattern')
		pattern = xlwt.Pattern()
		pattern.pattern = pattern.SOLID_PATTERN
		pattern.pattern_fore_colour = temp_bg_color_index

		style.pattern = pattern
		style.font = font
		style.borders = borders
		return style

	def create_excel(self,sheet_name):
		workbook = xlwt.Workbook() #创建工作簿
		workbook.add_sheet(u'所有',cell_overwrite_ok=True) #创建sheet
		workbook.save(self.full_path) #保存文件
		print (u'\t\tExcel创建完毕!')
        
	def open_excel(self):
		workbook = xlrd.open_workbook(self.full_path)
		return copy(workbook)
        
	#写入原始数据，如表格形式等
	def write_original_data(self):
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		row0 = [u'单词',u'出现频率',u'音标',u'中文翻译',u'书中例句']
			#生成第一行
		for i in range(1,len(row0)+1):
			sheet.write(1,i,row0[i-1],self.title_style)
		workbook.save(self.full_path)

		
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
			word_frequency = self.txt_dealer.get_map_word_frequency()
			for (key,value) in word_frequency.items():
				if str.isdigit(key.encode('gbk')) or len(key)<2 or check_contain_num(key) :
					pass
				else:
					sheet.write(row,1,key)
					sheet.write(row,2,value)
					row += 1
			workbook.save(self.full_path)#该操作耗时太长
		#print ('write_word_frequency\n')

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

	def search_sentence_by_word(self, word):
		#需要提前去除中文空格
		try:
		# for sentence in self.sentences:
		# #list_word = list[s.lower() for s in re.findall(r"\w+", sentence)]
		# if word in [s.lower() for s in re.findall(r"\w+", sentence)]:
		# return sentence
		# return 'NULL'
			ret = re.search(r'([^.]*{}[^.]*\.)'.format(word), self.data, re.IGNORECASE)
		# for i, r in enumerate(sentence, 1):
		# s = r
		# return r+'.'
			if ret:
				rst = ret.group()
			#有几个无符号空白不知道怎么去除
				repl = re.compile(r'[\s]+')
				out = re.sub(repl, ' ', rst)
				#out = out.strp()
				#out =out.replace('\n',' ')
			else:
				out = 'NULL'
			return out
		except Exception as e:
			print (word)

	#def search_sentence_by_word(self,word):
	#	sentence = re.findall(r'[^.]*?{}[^.]*?.'.format(word), self.data)
	#	for i, r in enumerate(sentence, 1):
	#		s = r
	#		return r+'.'
	#	return 'NULL'
		
	def get_map_word_frequency(self):
		return self.map_word_frequency
		
	def write_in_txt(self):
		self.fopen = open(self.path+".txt",'w')
		for word in self.map_word_frequency :
			self.fopen.write(str(word)+'\n')
		self.fopen.close()
		
class net_work:
	def __init__(self,config):
		self.config = config
		#path = self.config.get_conf(r'tools', r'webdriver_path')
		#self.web_driver =webdriver.Chrome(executable_path = path)
		#self.web_driver = self.init_webdriver()

	def init_webdriver(self):
		path =self.config.get_conf(r'tools',r'webdriver_path')
		return webdriver.Chrome(executable_path = path)
	def youdao_down_load(self,word):
		try:
			path = r'./https/youdao/{}.html'.format(word)
			# 支持断网续传
			if os.path.exists(path) or word == '':
				# print('{} is already exists'.format(path))
				pass
			else:
				url = r'http://dict.youdao.com/w/eng/{}/#keyfrom=dict2.index'.format(word)
				headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
				req = urllib2.Request(url, headers=headers)
				content = urllib2.urlopen(req).read()
				fp = open(path, 'w')
				fp.write(content)
				fp.close()
		except Exception as e:
			print (e)
	def https_down_load(self,word):
		try:
			path = r'./https/{}.html'.format(word)
			# 支持断网续传
			if os.path.exists(path) or word == '':
				# print('{} is already exists'.format(path))
				pass
			else:
				url = r'https://www.youdict.com/w/{}'.format(word)
				ssl._create_default_https_context = ssl._create_unverified_context
				headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
				req = urllib2.Request(url, headers=headers)
				content = urllib2.urlopen(req).read()
				fp = open(path, 'w')
				fp.write(content)
				fp.close()
		except Exception as e:
			print(e)
	def urlretrieve_down_load(self,word, num):
		try:
			path = r'./audios/{}.mpeg'.format(word)
			if os.path.exists(path):
				# print('{} is already exists'.format(path))
				pass
			else:
				url = r'https://dict.youdao.com/dictvoice?audio={}&type=2'.format(word)
				# time.sleep(1)
				urllib.request.urlretrieve(url, path)
		except urllib.error as e:
			print (e)
			num += 1
			time.sleep(num)
			self.urlretrieve_down_load(word, num)
	def get_data_by_https(self,word):
		ssl._create_default_https_context = ssl._create_unverified_context
		url = self.config.get_conf(r'web',r'path')
		url += word
		headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
		# 支持断网续传
		req = urllib2.Request(url, headers=headers)
		content = urllib2.urlopen(req).read()
		#print content
		if isinstance(content, str):
			pass
		else:
			content = content.decode('utf-8')
		html = HTML.fromstring(content)
		# 美式音标
		temp_soundmark = html.xpath(self.config.get_conf(r'web', r'path_soundmark'))
		if len(temp_soundmark) == 0:
			soundmark = r'NULL'
		else:
			soundmark = temp_soundmark[0].xpath(r'string()')
		# 翻译
		temp_translate = html.xpath(self.config.get_conf(r'web', r'path_translate'))
		if len(temp_translate) == 0:
			translate = r'NULL'
		else:
			translate = temp_translate[0].xpath(r'string(.)')
		# 词根
		temp_root = html.xpath(self.config.get_conf(r'web', r'path_root'))
		if len(temp_root) == 0:
			root = r'NULL'
		else:
			root = temp_root[0].xpath(r'string(.)')
		# 词源
		temp_etymology = html.xpath(self.config.get_conf(r'web', r'path_etymology'))
		if len(temp_etymology) == 0:
			etymology = r'NULL'
		else:
			etymology = temp_etymology[0].xpath(r'string(.)')

		data_list = [soundmark, translate, root, etymology]
		#print data_list
		#time.sleep(6)
		return data_list

	def get_data_by_selenium(self,word):
		path_driver = self.config.get_conf(r'tools', r'webdriver_path')
		web_driver =  webdriver.Chrome(executable_path=path_driver)
		path =self.config.get_conf(r'web',r'path').format(word)
		web_driver.get(path)
		#页面源码
		html = HTML.fromstring(web_driver.page_source)
		#美式发音 
		#pronounce = html.xpath()
		#美式音标
		temp_soundmark = html.xpath(self.config.get_conf(r'web',r'path_soundmark'))
		if len(temp_soundmark) == 0:
			soundmark =r'NULL'
		else:
			soundmark = temp_soundmark[0].xpath(r'string()')
		#翻译
		temp_translate= html.xpath(self.config.get_conf(r'web',r'path_translate'))
		if len(temp_translate) == 0:
			translate =r'NULL'
		else:
			translate = temp_translate[0].xpath(r'string(.)')
		#词根
		temp_root =  html.xpath(self.config.get_conf(r'web',r'path_root'))
		if len(temp_root) == 0:
			root =r'NULL'
		else:
			root = temp_root[0].xpath(r'string(.)')
		#词源
		temp_etymology = html.xpath(self.config.get_conf(r'web',r'path_etymology'))
		if len(temp_etymology) == 0:
			etymology =r'NULL'
		else:
			etymology = temp_etymology[0].xpath(r'string(.)')
		
		data_list = [soundmark,translate,root,etymology]
		#web_driver.quit()
		web_driver.close()
		return data_list

class config_dealer:
	def __init__(self):
		self.path = r'config.ini'
		self.init_config()
		self.parser = ConfigParser.ConfigParser()
		#self.parser = self.init_configparser()
	
	
	def init_configparser(self):
		tmp = ConfigParser.ConfigParser()
		return tmp.readfp(open(self.path))
	
	def get_conf(self,section,item):
		self.parser.readfp(open(self.path))
		return self.parser.get(section,item)
		
	def read_config(self):
		open(self.path)
	#判断存在
	def init_config(self):
		if os.path.exists(self.path):
			pass
		else:
			self.create_origin()
	#不存在就创建
	def create_origin(self):
		file =open(self.path,'w')
		br ='\n'
		
		#工具配置
		file.writelines(r'[book]'+br)
		file.writelines(r'path = steve jobs.txt'+br)
		file.writelines(r'[tools]'+br)
		file.writelines(r'webdriver_path = C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe'+br)
		#url配置
		file.writelines(r'[web]'+br)
		file.writelines(r'spider_rest_time = 4'+br)
		file.writelines(r'#以下为单词页面的url格式，以{}表示变动的位置'+br)
		file.writelines(r'#如：https://www.youdict.com/w/{}  默认访问youdict'+br)
		file.writelines(r'path = https://www.youdict.com/w/{}'+br)
		file.writelines(r'path_soundmark = //*[@id="yd-word-pron"]/span[2]'+br)
		file.writelines(r'path_translate = //*[@id="yd-word-meaning"]'+br)
		file.writelines(r'path_root = //*[@id="yd-ciyuan"]/p'+br)
		file.writelines(r'path_etymology = //*[@id="yd-etym"]'+br)
		file.writelines(r'path_audio =https://www.youdict.com/w/{}' + br)
		# 音频
		file.writelines(r'[audio]' + br)
		file.writelines(r'path =./audios/' + br)
		file.writelines(r'color =13' + br)
		# 日志
		file.writelines(r'[log]' + br)
		file.writelines(r'path = ./log/' + br)
		#excel
		file.writelines(r'[excel]' + br)
		file.writelines(r'#表头内容' + br)
		file.writelines(r'title_word = 单词' + br)
		file.writelines(r'title_prononciation = 音标' + br)
		file.writelines(r'title_frequency = 出现频率' + br)
		file.writelines(r'title_translation = 中文翻译' + br)
		file.writelines(r'title_root = 词根' + br)
		file.writelines(r'title_original = 词源' + br)
		file.writelines(r'title_example = 书中例句' + br)
		file.writelines(r'#字体样式' + br)
		file.writelines(r'style_font_color_index = 0' + br)
		file.writelines(r'style_font_name = 微软雅黑' + br)
		file.writelines(r'style_font_bold = True' + br)
		file.writelines(r'style_font_height = 6' + br)
		file.writelines(r'#框线样式' + br)
		file.writelines(r'style_border_left = 2' + br)
		file.writelines(r'style_border_right = 2' + br)
		file.writelines(r'style_border_top = 2' + br)
		file.writelines(r'style_border_bottom = 2' + br)
		file.writelines(r'#背景样式' + br)
		file.writelines(r'style_bg_color_index = 5' + br)
		#file.writelines(r'style_bg_pattern = Pattern.SOLID_PATTERN' + br)
		file.writelines(r'#单元格尺寸' + br)
		file.writelines(r'width_word = 50' + br)
		file.writelines(r'width_prononciation = 100' + br)
		file.writelines(r'width_frequency = 150' + br)
		file.writelines(r'width_translation = 200' + br)
		file.writelines(r'width_root = 250' + br)
		file.writelines(r'width_original = 300' + br)
		file.writelines(r'width_example = 350' + br)

		file.close()
		
class System_Thing:
    def __init__(self):
        self.path = self.get_path()
        
    def get_path(self):
        tail = str(r'/')
        return sys.path[0]+tail

    def file_is_exist(self,filename):
        return os.path.exists(self.path+filename)

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