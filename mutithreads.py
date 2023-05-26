# -*- coding:utf-8 -*-
import sys,os
import urllib2
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from threading import Thread
import ssl
import re
import io
# def muti_thread():
	# p = Pool(4)
	# for i in range(5):
		# p.apply_async(self.write_word_meaning(), args=(i,))
	# print ("等待所有子进程执行完毕...")
	# p.close()
	# p.join()
	
def cycle(map):
	for (key,value) in map.items():
		https_down_load(key)
		
def https_down_load(word):
	try:
		path = r'./https/{}.html'.format(word)
		# 支持断网续传
		if os.path.exists(path) or word =='':
			# print('{} is already exists'.format(path))
			pass
		else:
			url = r'https://www.youdict.com/w/{}'.format(word)
			ssl._create_default_https_context = ssl._create_unverified_context
			headers = {'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
			req = urllib2.Request(url, headers=headers)
			content = urllib2.urlopen(req).read()
			fp =open(path,'w')
			fp.write(content)
			fp.close()
	except Exception,e:
		print e
		s =1
		
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
	
if __name__ =='__main__':
	dealer = text_dealer(r'Steve Jobs.txt')
	map = dealer.get_map_word_frequency
	list =[]
	s =0
	num = len(map)/10
	for i in range(0,10)
		list.append(map[s:i*num] )
		s = i*num
	list.append(9*num:)	
	
	p = Pool(10)
	for t in range(10):
		p.apply_async(cycle, args=(list[t],))
	print ("等待所有子进程执行完毕...")
	p.close()
	p.join()