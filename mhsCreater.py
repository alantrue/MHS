# -*- coding: utf-8 -*-

import os
import csv
import shutil
import zipfile

'''
files = os.listdir("C:\\Users\\alantrue\\Desktop\\work")
for filename in files:
	print(filename)
'''
workPath = u'C:/Users/alantrue/Desktop/work/'
tempPath = u'C:/Users/alantrue/Desktop/work/temp/'
files = (u'無口語紀錄表.docx', u'無口語評估表.docx', u'詞彙簡單句紀錄表.docx', u'詞彙簡單句評估表.docx', u'構音紀錄表.docx', u'構音評估表.docx', u'複雜句故事敘述紀錄表.docx', u'複雜句故事敘述評估表.docx')


def parseInfo():
	#解析出0病歷號, 1姓名, 2生日, 3初診日, 4性別, 5階段, 6紀錄
	csvfile = open('C:\\Users\\alantrue\\Desktop\\work\\info.csv', 'rb')

	i = 0
	for row in csv.reader(csvfile, delimiter=',', quotechar='"'):
		fileSrc1 = u''
		fileSrc2 = u''
		fileDes1 = u'治療表.zip'.encode('big5')
		fileDes2 =  u'評估表.zip'.encode('big5')
		#依據階段複製範本, 並命名病歷號+姓名+(語言or構音)+(治療, 評估)表
   		if row[5] == u'無口語'.encode('big5'):
   			#複製 無口語紀錄表.docx, 無口語評估表.docx
   			fileSrc1 = files[0]
   			fileSrc2 = files[1]
   			fileDes1 = u'語言'.encode('big5') + fileDes1
   			fileDes2 = u'語言'.encode('big5') + fileDes2

		elif row[5] == u'簡單句'.encode('big5'):
			#複製 詞彙簡單句紀錄表.docx, 詞彙簡單句評估表.docx
			fileSrc1 = files[2]
   			fileSrc2 = files[3]
   			fileDes1 = u'語言'.encode('big5') + fileDes1
   			fileDes2 = u'語言'.encode('big5') + fileDes2

		elif row[5] == u'構音'.encode('big5'):
			#複製 構音紀錄表.docx, 構音評估表.docx
			fileSrc1 = files[4]
   			fileSrc2 = files[5]
   			fileDes1 = u'構音'.encode('big5') + fileDes1
   			fileDes2 = u'構音'.encode('big5') + fileDes2

		elif row[5] == u'複雜句'.encode('big5'):
			#複製 複雜句故事敘述紀錄表.docx, 複雜句故事敘述評估表.docx
			fileSrc1 = files[6]
   			fileSrc2 = files[7]
   			fileDes1 = u'語言'.encode('big5') + fileDes1
   			fileDes2 = u'語言'.encode('big5') + fileDes2

		fileDes1 = row[0] + row[1] + fileDes1
		fileDes2 = row[0] + row[1] + fileDes2

		if fileSrc1 != u'' and fileSrc2 != u'':

			fileSrc1 = workPath + fileSrc1
			fileSrc2 = workPath + fileSrc2
			fileDes1 = tempPath.encode('big5') + fileDes1
			fileDes2 = tempPath.encode('big5') + fileDes2
			print row[0], row[1], row[2], row[3], row[4], row[5], row[6]
			print i, fileSrc1, fileSrc2, fileDes1, fileDes2
			#建立拷貝zip
			shutil.copyfile(fileSrc1, fileDes1.decode('big5'))
			shutil.copyfile(fileSrc2, fileDes2.decode('big5'))
		#解壓縮
			zf = zipfile.ZipFile(fileDes1.decode('big5'), 'r')
			data = zf.read('word/document.xml')
			data = data.decode('utf8')
		#填入資料
			#病歷號碼
			data = data.replace(u'123456', row[0].decode('big5'))
			#名字
			data = data.replace(u'我的名字', row[1].decode('big5'))
			#生日
			data = data.replace(u'birth', row[2].decode('big5'))
			#初診日期
			#男女

			print data
			#data.encode('utf8')
			#zf.write(data,'word/document.xml')
		#壓縮回zip
		#改回docx

		i = i + 1
		if i > 1:
			break

parseInfo()