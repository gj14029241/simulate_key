# coding = utf-8

import win32com.client
import json

conf = open("wpsconf.txt",encoding='utf-8').read()
conf = json.loads(conf)

app = win32com.client.Dispatch('ket.Application')
app.Sheets(1).Activate()

for hengpai in conf["hengpai"]:
	col_char = hengpai["col_char"]
	t_row = hengpai["t_row"]
	xin_no = 1
	comment_rg = app.Sheets(1).Range("%s%d" % (col_char,t_row))
	for col in hengpai["text"]:
		for key in col:
			l = key.split("-")
			base_text = col[key]
			for i in range(int(l[0]), int(l[1])+1):
				text = "%d-%d至%s" % (xin_no, xin_no + 11, base_text)
				xin_no += 12
				comment_rg.AddComment().Text(text)
				color_rg = comment_rg.Offset(0,1)
				for j in range(1,13):
					color_rg.Interior.Color = 0x32d599
					color_rg = color_rg.Offset(0,1)
				comment_rg = comment_rg.Offset(1, 3)
			break


for shupai in conf["shupai"]:
	col_char = shupai["col_char"]
	t_row = shupai["t_row"]
	xin_no = 1
	for row in shupai["text"]:
		for key in row:
			l = key.split("-")
			base_text = row[key]
			for i in range(int(l[0]), int(l[1])+1):
				text = "%d-%d至%s" % (xin_no, xin_no + 11, base_text)
				xin_no += 12
				color_rg = app.Sheets(1).Range("%s%d" % (col_char,t_row))
				color_rg.Value = text
				color_rg = color_rg.Previous
				color_rg = color_rg.Previous
				for j in range(1,13):
					color_rg = color_rg.Previous
					color_rg.Interior.Color = 0x32d599
				t_row += 2
			break



