# -*- coding:utf-8 -*-
import sys,os
import urllib.request as urllib2
from multiprocessing import Pool
from scripts import TextDealer
#import urllib3
#import requests
#from requests.packages.urllib3.exceptions import InsecureRequestWarning
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
	except Exception as e:
		print (e)
		s =1


if __name__ == '__main__':
	dealer = TextDealer(r'Steve Jobs.txt')
	mapWd = dealer.get_map_word_frequency
	listWd = []
	s = 0
	num = len(mapWd)/10
	for i in range(0, 10):
		listWd.append(mapWd[s:i*num] )
		s = i*num
	listWd.append(mapWd[9*num:])
	
	p = Pool(10)
	for t in range(10):
		p.apply_async(cycle, args=(listWd[t],))
	print("等待所有子进程执行完毕...")
	p.close()
	p.join()
