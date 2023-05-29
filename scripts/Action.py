from scripts import ConfigDealer
from scripts import TextDealer
from scripts import Network
from scripts import ExcelDealer

class Action:
	def __init__(self, path):
		try:
			print(u'配置模块初始化...')
			self.config_dealer = ConfigDealer()
			# mark
			self.path = path
			# self.path = self.config_dealer.get_conf(r'book',r'path')
			print(u'文本模块初始化...')
			self.txt_dealer = TextDealer(self.path)
			print(u'网络模块初始化...')
			self.net_work_dealer = Network(self.config_dealer)
			print(u'Excel模块初始化...')
			self.excel_dealer = ExcelDealer(self.config_dealer, self.txt_dealer, self.net_work_dealer, self.path)
		except Exception as error_info:
			print(u'【action_init ERROR】:{}'.format(error_info))
			input("Enter enter key to continue...")

	def write_info_into_excel(self):
		try:
			print(u'执行音频下载......')
			# self.excel_dealer.down_load_audio()

			# get = int(raw_input("Insert the audio?('0'for Yes,other for no)"))

			# if get ==0:
			print(u'执行音频插入......')
			# self.excel_dealer.insert_audio()

			# print u'执行信息写入......'
			# #govent = [gevent.spawn(self.excel_dealer.write_word_meaning) for i in range(0,9)]
			# #gevent.joinall(g)
			# ts = [Thread(target=self.excel_dealer.write_word_meaning,args=(i,)) for i in range(0,9)]
			# for t in ts:
			# t.start()
			# for t in ts:
			# t.join()
			self.excel_dealer.write_word_meaning()

			print(u'执行信息优化')
			# self.excel_dealer.optimisation()

			print(u'信息写入完毕！')
		except Exception as error_info:
			print(u'【action_init ERROR】:{}'.format(error_info))
			input("Enter enter key to continue...")
