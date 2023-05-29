import os,time
from selenium import webdriver
import urllib.request as urllib2
import ssl
import urllib
import lxml.html as HTML


class Network:
	def __init__(self ,config):
		self.config = config

	# path = self.config.get_conf(r'tools', r'webdriver_path')
	# self.web_driver =webdriver.Chrome(executable_path = path)
	# self.web_driver = self.init_webdriver()

	def init_webdriver(self):
		path =self.config.get_conf(r'tools', r'webdriver_path')
		return webdriver.Chrome(executable_path=path)

	def youdao_down_load(self, word):
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
			print(e)

	def https_down_load(self, word):
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

	def urlretrieve_down_load(self, word, num):
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
			print(e)
			num += 1
			time.sleep(num)
			self.urlretrieve_down_load(word, num)

	def get_data_by_https(self, word):
		ssl._create_default_https_context = ssl._create_unverified_context
		url = self.config.get_conf(r'web', r'path')
		url += word
		headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
		# 支持断网续传
		req = urllib2.Request(url, headers=headers)
		content = urllib2.urlopen(req).read()
		# print content
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
		# print data_list
		# time.sleep(6)
		return data_list

	def get_data_by_selenium(self, word):
		path_driver = self.config.get_conf(r'tools', r'webdriver_path')
		web_driver = webdriver.Chrome(executable_path=path_driver)
		path = self.config.get_conf(r'web', r'path').format(word)
		web_driver.get(path)
		# 页面源码
		html = HTML.fromstring(web_driver.page_source)
		# 美式发音
		# pronounce = html.xpath()
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
		# web_driver.quit()
		web_driver.close()
		return data_list
