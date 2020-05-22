#!usr/bin/env python
# -*- coding: utf-8 -*-

import os
import json
import re
import xlwt


class JS2Excel():

	LANG_MAP = {
		"zh": "中文",
		"en": "英文",
		"ja": "日文",
		"pt": "葡萄牙语",
		"es": "西班牙",
		"fr": "法语",
		"ko": "韩语",
		"ru": "俄语",
		"th": "泰语",
		"id": "印尼语",
		"tr": "土耳其语",
		"vi": "越南语",
		"it": "意大利语",
		"el": "希腊语",
		"ar": "阿拉伯语",
		"da": "丹麦语",
		"fa": "波斯语",
		"fi": "芬兰语",
		"pi": "波兰语",
		"sv-SE": "瑞典语",
		"ua": "乌克兰语",
		"sr": "塞尔维亚语",
		"ro": "罗马尼亚语",
		"hu": "匈牙利语",
	}

	JS_MODE = {
		"es": re.compile("export\s+default\s*"),
		"common": re.compile("module.exports\s*=\s*"),
		"normal": re.compile("")
	}

	def __init__(self, input_dir, output_file, js_mode='es'):
		# 输出的Excel路径
		self.outputDir = output_file
		# 前端js文件的模块类型commjs, es, normal(非模块化)
		self.js_mode = js_mode
		# 前端语言配置表的文件夹
		self.parse_js_to_excel(input_dir)


	def parse_js_to_excel(self, path):
		files = os.listdir(path)
		wb = xlwt.Workbook()
		ws = wb.add_sheet('Sheet1', cell_overwrite_ok=True)
		font = xlwt.Font()
		# 设置字体为红色
		font.colour_index=xlwt.Style.colour_map['red']
		style = xlwt.XFStyle()
		style.font = font
		ws.row(0).height_mismatch = True
		ws.row(0).height = 1000
		alignment = xlwt.Alignment()
		alignment.vert = xlwt.Alignment.VERT_CENTER
		style.alignment = alignment
		ws.write_merge(0,0,0,100, '第一列和第二行的数据为自动生成，请不要更改', style)

		'''
			Excel中，规定从第二行开始为翻译的内容，第一行为说明区域，第二行开始为翻译内容的表头：
			第一列表为前端对应的键，从第二列开始对应语言的翻译
			如：
			key    中文  英文   日语
			hello  你好  hello  歓迎する
		'''
		# 第二行第一列的表头
		ws.write(1,0,'key')
		reg = JS2Excel.JS_MODE[self.js_mode]
		# 用来解决不同语言配置键顺序不一致、键不对应和嵌套结构写入Excel行号递增等问题
		rows = []
		for i in range(0, len(files)):
			"""
				通过文件名获取到相应的语言配置信息
				要求文件夹目录结构为
				-[self.outputDir] # 选择的文件夹
			  --zh.js
			  --en.js
			  --ja.js
			  ... 
			"""
			f = files[i]
			pos = f.rindex('.')
			lang = f[:pos]

			# 只支持LANG_MAP配置的语言信息
			if lang not in JS2Excel.LANG_MAP:
				continue
			ws.write(1, i + 1, JS2Excel.LANG_MAP[lang])

			with open(os.path.join(path,f), 'r', encoding='utf-8') as f:
				json_str = f.read()
				# 去掉前端模块中的export default字符串
				try:
					json_str = re.sub(reg, '', json_str)
					json_dic = json.loads(json_str, encoding="UTF-8")
				except:
					print("%s文件内容格式错误,请导出为一个json结构的数据" % f)
				else:
					self.write_to_excel(json_dic, ws, rows, i + 1, '')

		# 导出保存为Excel
		wb.save(self.outputDir)
		print('done')

	def write_to_excel(self, data, ws, rows, col, key):
		'''
			Excel中，规定从第二行开始为翻译的内容，第一行为说明区域，第二行开始为翻译内容的表头：
			第一列表为前端对应的键，从第二列开始对应语言的翻译
			如：
			key    中文  英文   日语
			hello  你好  hello  歓迎する
		'''
		# 从第三行开始写翻译内容的key-value形式
		for k, val in data.items():
			strs = (k if key == '' else "%s@%s" % (key, k))
			if isinstance(val, dict):
				# 处理嵌套的结构
				self.write_to_excel(val, ws, rows, col, strs)
			else:
				# 如果在不同语言配置表中的顺序不一致时，会导致键错乱的问题，所以要对应到相应的键
				# row列表中键对应的下标加上2(因为是从第2行开始)即为该键在Excel中的行数
				r = 0
				if strs not in rows:
					r = len(rows) + 2
					rows.append(strs)
					ws.write(r, 0, strs)
				else:
					r = rows.index(strs) + 2
				ws.write(r, col, val)


def main():
	JS2Excel()

if __name__ == "__main__":
	main()




