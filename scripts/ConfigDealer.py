
import configparser
import os


class ConfigDealer:
	def __init__(self):
		self.path = r'config.ini'
		self.init_config()
		self.parser = configparser.ConfigParser()

	# self.parser = self.init_configparser()
	def init_configparser(self):
		tmp = configparser.ConfigParser()
		return tmp.read_file(open(self.path))

	def get_conf(self, section, item):
		self.parser.read_file(open(self.path))
		return self.parser.get(section, item)

	def read_config(self):
		open(self.path)

	# 判断存在
	def init_config(self):
		if os.path.exists(self.path):
			pass
		else:
			self.create_origin()

	# 不存在就创建
	def create_origin(self):
		file = open(self.path, 'w')
		br = '\n'

		# 工具配置
		file.writelines(r'[book]' + br)
		file.writelines(r'path = steve jobs.txt' + br)
		file.writelines(r'[tools]' + br)
		file.writelines(r'webdriver_path = C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe' + br)
		# url配置
		file.writelines(r'[web]' + br)
		file.writelines(r'spider_rest_time = 4' + br)
		file.writelines(r'#以下为单词页面的url格式，以{}表示变动的位置' + br)
		file.writelines(r'#如：https://www.youdict.com/w/{}  默认访问youdict' + br)
		file.writelines(r'path = https://www.youdict.com/w/{}' + br)
		file.writelines(r'path_soundmark = //*[@id="yd-word-pron"]/span[2]' + br)
		file.writelines(r'path_translate = //*[@id="yd-word-meaning"]' + br)
		file.writelines(r'path_root = //*[@id="yd-ciyuan"]/p' + br)
		file.writelines(r'path_etymology = //*[@id="yd-etym"]' + br)
		file.writelines(r'path_audio =https://www.youdict.com/w/{}' + br)
		# 音频
		file.writelines(r'[audio]' + br)
		file.writelines(r'path =./audios/' + br)
		file.writelines(r'color =13' + br)
		# 日志
		file.writelines(r'[log]' + br)
		file.writelines(r'path = ./log/' + br)
		# excel
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
		# file.writelines(r'style_bg_pattern = Pattern.SOLID_PATTERN' + br)
		file.writelines(r'#单元格尺寸' + br)
		file.writelines(r'width_word = 50' + br)
		file.writelines(r'width_prononciation = 100' + br)
		file.writelines(r'width_frequency = 150' + br)
		file.writelines(r'width_translation = 200' + br)
		file.writelines(r'width_root = 250' + br)
		file.writelines(r'width_original = 300' + br)
		file.writelines(r'width_example = 350' + br)

		file.close()
