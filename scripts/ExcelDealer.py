import os,time
from scripts import SystemThing
import xlwt
import xlrd
import lxml.html as HTML
from xlutils.copy import copy

def check_contain_num(check_str):
    flag = False
    for ch in check_str.decode('utf-8'):
        if u'9' >= ch and ch >= u'0':
            flag =  True
    return flag

class ExcelDearler:
	def __init__(self, config, txt_dealer, net_work, path):
		self.conf = config
		self.txt_dealer = txt_dealer
		self.net_work = net_work

		self.file_name = ''
		self.full_path = ''
		self.init_config(path)

		self.title_style = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: bold on;')
		self.content_styleA = self.set_title_style(5)
		self.content_styleB = self.set_title_style(4)
		print(u'\t表格样式设置完毕！')
		self.write_word_frequency()
		print(u'\t原始表格数据写入完毕！')
		self.select()
		print(u'根据熟词库将单词筛选完毕！')

	# self.write_all()
	# self.write_word_meaning()

	def select(self):
		# mark 新加入的根据历史记录过滤熟词
		root_path = "output/"
		file_lst = os.listdir(root_path)
		for file in file_lst:
			full_path = root_path + file
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

	def insert_youdao_mean_from_loacl(self, word):
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

	def insert_mean_from_local(self, word):
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

	def insert_audio(self):  # 重复处理后会丢失超链接和背景色
		workbookr = xlrd.open_workbook(self.file_name, formatting_info=True)
		sheetr = workbookr.sheet_by_index(0)  # sheet索引从0开始
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		# 区分超链接背景色,重复写入超链接会造成空
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
				row += 1
				continue
			prepath = self.conf.get_conf(r'audio', 'path')
			path = prepath + word + '.mpeg'
			if os.path.exists(path):
				hyper = u'HYPERLINK("{}";"{}")'.format(path, word)
				sheet.write(row, 1, xlwt.Formula(hyper), style)
				counter += 1
			row += 1
		self.save(workbook)
		print(u'插入{}条音频，共{}条数据'.format(counter, sheetr.nrows))

	def init_config(self, path):
		self.file_name = path + r'.xls'
		self.full_path = self.file_name

	def save(self, workbook):
		try:
			# mark
			root_path = 'output/'
			workbook.save(root_path + self.full_path)
		except:
			print('please close the excle')
			time.sleep(4)
			self.save(workbook)

	def write_word_meaning(self):
		workbookr = xlrd.open_workbook(self.file_name)
		sheetr = workbookr.sheet_by_index(0)  # sheet索引从0开始

		print(u'开始写入网页数据....（该处理时间较长，停一会儿就可以看效果,随时可以关闭程序）')
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		for row in range(2, sheetr.nrows):
			value = sheetr.cell(row, 4).value
			# 二次传输时自动识别弥补，不覆盖已有
			# 单元格内容不为空
			if value != '':
				# row += 10
				continue
			word = sheetr.cell(row, 1).value.encode('utf-8')
			if word == '':
				continue
			# time.sleep(3)
			# data_list = self.net_work.get_data_by_selenium(word)
			# data_list = self.net_work.get_data_by_https(word)
			data_list = self.insert_youdao_mean_from_loacl(word)
			# data_list = self.insert_mean_from_local(word)
			sentence = self.txt_dealer.search_sentence_by_word(word)
			for cnt in range(0, 2):
				info = data_list[cnt]
				sheet.write(row, cnt + 3, info)
			sheet.write(row, 7, sentence)
		self.save(workbook)
		# row+=1
		# if row % 10==0:
		# self.save(workbook)
		# print u'已处理{}条数据'.format(row-1)
		print('All the meanning be written\n')

	def set_title_style(self, bg_clolor=False):
		# 初始化样式
		style = xlwt.XFStyle()
		# 创建字体
		temp_font_color_index = self.conf.get_conf(r'excel', 'style_font_color_index')
		temp_font_name = self.conf.get_conf(r'excel', 'style_font_name')
		temp_font_bold = self.conf.get_conf(r'excel', 'style_font_bold')
		temp_font_height = self.conf.get_conf(r'excel', 'style_font_height')
		font = xlwt.Font()
		font.name = temp_font_name
		font.bold = temp_font_bold
		font.color_index = temp_font_color_index
		font.height = temp_font_height
		# 创建下框线
		temp_border_left = self.conf.get_conf(r'excel', 'style_border_left')
		temp_border_right = self.conf.get_conf(r'excel', 'style_border_right')
		temp_border_top = self.conf.get_conf(r'excel', 'style_border_top')
		temp_border_bottom = self.conf.get_conf(r'excel', 'style_border_bottom')
		borders = xlwt.Borders()
		borders.left = temp_border_left
		borders.right = temp_border_right
		borders.top = temp_border_top
		borders.bottom = temp_border_bottom

		# 创建背景颜色
		if bg_clolor != False:
			temp_bg_color_index = bg_clolor
		else:
			temp_bg_color_index = self.conf.get_conf(r'excel', 'style_bg_color_index')
		# temp_bg_pattern = self.conf.get_conf(r'excel', 'style_bg_pattern')
		pattern = xlwt.Pattern()
		pattern.pattern = pattern.SOLID_PATTERN
		pattern.pattern_fore_colour = temp_bg_color_index

		style.pattern = pattern
		style.font = font
		style.borders = borders
		return style

	def create_excel(self, sheet_name):
		workbook = xlwt.Workbook()  # 创建工作簿
		workbook.add_sheet(u'所有', cell_overwrite_ok=True)  # 创建sheet
		workbook.save(self.full_path)  # 保存文件
		print(u'\t\tExcel创建完毕!')

	def open_excel(self):
		workbook = xlrd.open_workbook(self.full_path)
		return copy(workbook)

	# 写入原始数据，如表格形式等
	def write_original_data(self):
		workbook = self.open_excel()
		sheet = workbook.get_sheet(0)
		row0 = [u'单词', u'出现频率', u'音标', u'中文翻译', u'书中例句']
		# 生成第一行
		for i in range(1, len(row0) + 1):
			sheet.write(1, i, row0[i - 1], self.title_style)
		workbook.save(self.full_path)

	# 写excel
	def write_word_frequency(self):
		syst = SystemThing.SystemThing()
		if syst.file_is_exist(self.full_path):
			print('{} is already exists'.format(self.full_path))
			pass
		else:
			self.create_excel('sheet1')
			self.write_original_data()
			workbook = self.open_excel()
			sheet = workbook.get_sheet(0)
			# 填写数据word_frequency
			row = 2
			word_frequency = self.txt_dealer.get_map_word_frequency()
			for (key, value) in word_frequency.items():
				if str.isdigit(key.encode('gbk')) or len(key) < 2 or check_contain_num(key):
					pass
				else:
					sheet.write(row, 1, key)
					sheet.write(row, 2, value)
					row += 1
			workbook.save(self.full_path)  # 该操作耗时太长
# print ('write_word_frequency\n')
