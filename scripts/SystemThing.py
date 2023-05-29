import sys
import os


class SystemThing:
	def __init__(self):
		self.path = self.get_path()

	def get_path(self):
		tail = str(r'/')
		return sys.path[0] + tail

	def file_is_exist(self, filename):
		return os.path.exists(self.path + filename)
