# -*- coding: utf-8 -*-

import os
import csv

'''
files = os.listdir("C:\\Users\\alantrue\\Desktop\\work")
for filename in files:
	print(filename)
'''


def parseInfo():
	#解析出0病歷號, 1姓名, 2生日, 3初診日, 4性別, 5階段, 6紀錄
	csvfile = open('C:\\Users\\alantrue\\Desktop\\work\\info.csv', 'rb')
	for row in csv.reader(csvfile, delimiter=',', quotechar='"'):
		#print row[0], row[1], row[2], row[3], row[4], row[5], row[6]
		#依據階段複製範本, 並命名病歷號+姓名+(語言or構音)+(評估, 治療)表
   		if row[5] == u'無口語'.encode('big5'):
   			#複製 無口語紀錄表.docx, 無口語評估表.docx
		elif row[5] == u'簡單句'.encode('big5'):
			#複製 詞彙簡單句紀錄表.docx, 詞彙簡單句評估表.docx
		elif row[5] == u'複雜句'.encode('big5'):
			#複製 複雜句故事敘述評估表.docx, 複雜句故事描述紀錄表.docx
		elif row[5] == u'構音'.encode('big5'):
			#複製 構音紀錄表.docx, 構音評估表.docx

		#改成zip
		#解壓縮
		#填入資料
		#壓縮回zip
		#改回docx

parseInfo()