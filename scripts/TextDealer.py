import io
import re


class TextDealer:
	def __init__(self, path):
		self.path = path
		self.map_word_frequency = dict()
		# sentence = re.compile(r"\.+")
		with io.open(self.path, encoding="utf-8") as f:
			self.data = f.read()
			self.sentences = [s.lower() for s in re.findall(r"\.+", self.data)]
			self.words = [s.lower() for s in re.findall(r"\w+", self.data)]
		for word in self.words:
			self.map_word_frequency[word] = self.map_word_frequency.get(word, 0) + 1

	def search_sentence_by_word(self, word):
		# 需要提前去除中文空格
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
				# 有几个无符号空白不知道怎么去除
				repl = re.compile(r'[\s]+')
				out = re.sub(repl, ' ', rst)
			# out = out.strp()
			# out =out.replace('\n',' ')
			else:
				out = 'NULL'
			return out
		except Exception as e:
			print(word)

	# def search_sentence_by_word(self,word):
	#	sentence = re.findall(r'[^.]*?{}[^.]*?.'.format(word), self.data)
	#	for i, r in enumerate(sentence, 1):
	#		s = r
	#		return r+'.'
	#	return 'NULL'

	def get_map_word_frequency(self):
		return self.map_word_frequency

	def write_in_txt(self):
		self.fopen = open(self.path + ".txt", 'w')
		for word in self.map_word_frequency:
			self.fopen.write(str(word) + '\n')
		self.fopen.close()
