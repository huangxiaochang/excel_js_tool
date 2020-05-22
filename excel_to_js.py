#!usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import json
import os


class Excel2JS():

	LANG_MAP = {
		"中文": 'zh',
		"英文": 'en',
		"日文": 'ja',
		"葡萄牙语": "pt",
		"西班牙": "es",
		"法语": "fr",
		"韩语": "ko",
		"俄语": "ru",
		"泰语": "th",
		"印尼语": "id",
		"土耳其语": "tr",
		"越南语": "vi",
		"意大利语": "it",
		"希腊语": "el",
		"阿拉伯语": "ar",
		"丹麦语": "da",
		"波斯语": "fa",
		"芬兰语": "fi",
		"波兰语": "pi",
		"瑞典语": "sv-SE",
		"乌克兰语": "ua",
		"塞尔维亚语": "sr",
		"罗马尼亚语": "ro",
		"匈牙利语": "hu",
	}

	JS_MODE = {
		"es": "export default ",
		"common": "module.exports = ",
		"normal": ""
	}

	def __init__(self, input_file, output_dir, js_mode='es'):
		# 输出的语言配置js文件的目录
		self.outputDir = output_dir
		# 前端js文件的模块类型commjs, es, normal(非模块化)
		self.js_mode = js_mode
		# Excel翻译文件来源
		self.parse_excel_to_js(input_file)

	def parse_excel_to_js(self, path):
		excel = xlrd.open_workbook(path)
		table = excel.sheet_by_index(0)
		nrows = table.nrows
		ncols = table.ncols
		'''
			Excel中，规定从第二行开始为翻译的内容，第一行为说明区域，第二行开始为翻译内容的表头：
			第一列表为前端对应的键，从第二列开始对应语言的翻译
			如：
			key    中文  英文   日语
			hello  你好  hello  歓迎する
		'''
		# 获取语言的种类
		headers = table.row_values(1, start_colx=1)
		while '' in headers:
			headers.remove('')
		# 处理每一列
		for col in range(0, len(headers)):
			'''
			输出的文件结构为
			  -[self.outputDir] # 选择的文件夹
			  --zh.js
			  --en.js
			  --ja.js
			  ...
			'''
			# 只支持LANG_MAP配置的语言信息
			if headers[col] not in Excel2JS.LANG_MAP:
				continue

			filename = "%s/%s.js" % (self.outputDir, Excel2JS.LANG_MAP[headers[col]])
			json_dic = {}
			# 处理每一行
			for row in range(2, nrows):
				key = table.cell_value(row, 0)
				# 处理嵌套的结构，@字符分割为每一级
				keys = key.split("@")
				k = keys[-1]
				keys = keys[:-1]
				data = json_dic
				for i in range(0, len(keys)):
					if keys[i] not in data:
						data[keys[i]] = {}
					data = data[keys[i]]

				val = table.cell_value(row, col + 1)
				data[k] = val
			
			json_str = json.dumps(json_dic, ensure_ascii=False, indent=2, separators=(',', ':'))
			json_str = Excel2JS.JS_MODE[self.js_mode] + json_str
			if not os.path.exists(self.outputDir):
				os.makedirs(self.outputDir)

			# 创建js文件
			with open(filename, 'w', encoding='UTF-8') as f:
				f.write(json_str)
				f.close()

		print('Finished')


def main():
	Excel2JS()

if __name__ == "__main__":
	main()

