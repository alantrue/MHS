# -*- coding: utf-8 -*-

import os
import csv
import shutil
import zipfile
import tempfile
import datetime
import re
import random
import math

workPath = os.getcwd() + '/'
outputPath = workPath + u'病歷表/'
infoFile = workPath + 'info.txt'
files = (u'無口語紀錄表.docx', u'無口語評估表.docx', u'詞彙簡單句紀錄表.docx', u'詞彙簡單句評估表.docx', u'構音紀錄表.docx', u'構音評估表.docx', u'複雜句故事敘述紀錄表.docx', u'複雜句故事敘述評估表.docx')

def remove_from_zip(zipfname, *filenames):
    tempdir = tempfile.mkdtemp()
    try:
        tempname = os.path.join(tempdir, 'new.zip')
        with zipfile.ZipFile(zipfname, 'r') as zipread:
            with zipfile.ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.infolist():
                    if item.filename not in filenames:
                        data = zipread.read(item.filename)
                        zipwrite.writestr(item, data)
        shutil.move(tempname, zipfname)
    finally:
        shutil.rmtree(tempdir)

def replaceOccurence(str, search, replacement, index):
	split = str.split(search, index+1)
	if len(split) <= index+1:
		return str
	return search.join(split[:-1])+replacement+split[-1]

def randomProgress(p):
	for i in range(len(p)):
		if p[i] < 5:
			if p[i] == 1 and random.randint(0,99) < 25:
				p[i] += 1
			elif p[i] == 2 and random.randint(0,99) < 20:
				p[i] += 1
			elif p[i] == 3 and random.randint(0,99) < 10:
				p[i] += 1
			elif p[i] == 4 and random.randint(0,99) < 5:
				p[i] += 1

def removeFileInFolder(folder):
	filelist = [f for f in os.listdir(folder) if f.endswith(".docx")]
	for f in filelist:
		os.remove(folder + f)

def parseInfo():
	#解析出0病歷號, 1姓名, 2生日, 3初診日, 4性別, 5階段, 6紀錄
	csvfile = open(infoFile, 'r')

	i = 0
	for row in csv.reader(csvfile, delimiter='\t', quotechar='"'):
		if len(row) != 7:
			print u'有錯誤'.encode('big5')
			print u'info.txt 需要 [病歷號] [姓名] [生日] [初診日期] [性別] [階段] [紀錄] 等7個欄位'.encode('big5')
			break

		fileSrc1 = u''
		fileSrc2 = u''
		fileDes1 = u'治療表{0}.zip'.encode('utf8')
		fileDes2 =  u'評估表.zip'.encode('utf8')
		#依據階段複製範本, 並命名病歷號+姓名+(語言or構音)+(治療, 評估)表
   		if row[5] == u'無口語'.encode('utf8'):
   			#複製 無口語紀錄表.docx, 無口語評估表.docx
   			fileSrc1 = files[0]
   			fileSrc2 = files[1]
   			fileDes1 = u'語言'.encode('utf8') + fileDes1
   			fileDes2 = u'語言'.encode('utf8') + fileDes2

		elif row[5] == u'簡單句'.encode('utf8'):
			#複製 詞彙簡單句紀錄表.docx, 詞彙簡單句評估表.docx
			fileSrc1 = files[2]
   			fileSrc2 = files[3]
   			fileDes1 = u'語言'.encode('utf8') + fileDes1
   			fileDes2 = u'語言'.encode('utf8') + fileDes2

		elif row[5] == u'構音'.encode('utf8'):
			#複製 構音紀錄表.docx, 構音評估表.docx
			fileSrc1 = files[4]
   			fileSrc2 = files[5]
   			fileDes1 = u'構音'.encode('utf8') + fileDes1
   			fileDes2 = u'構音'.encode('utf8') + fileDes2

		elif row[5] == u'複雜句'.encode('utf8'):
			#複製 複雜句故事敘述紀錄表.docx, 複雜句故事敘述評估表.docx
			fileSrc1 = files[6]
   			fileSrc2 = files[7]
   			fileDes1 = u'語言'.encode('utf8') + fileDes1
   			fileDes2 = u'語言'.encode('utf8') + fileDes2

		fileDes1 = row[0] + row[1] + fileDes1
		fileDes2 = row[0] + row[1] + fileDes2

		if fileSrc1 != u''.encode('utf8') and fileSrc2 != u''.encode('utf8'):
			fileSrc1 = workPath + fileSrc1
			fileSrc2 = workPath + fileSrc2
			fileDes1 = outputPath.encode('utf8') + fileDes1
			fileDes2 = outputPath.encode('utf8') + fileDes2

			#建立拷貝zip
			p = [1, 1, 1, 1, 1, 1, 1, 1]
			try:
				needCount = math.ceil(len(row[6].split(',')) / 6.0)			
				if needCount < 2: #只需一份
					fileDes1 = fileDes1.format('')
					shutil.copyfile(fileSrc1, fileDes1.decode('utf8'))
					#產生紀錄表
					createMHS(row, fileDes1, p)
				else:
					for i in range(int(needCount)):
						fileDes1Temp = fileDes1.format(i+1)
						shutil.copyfile(fileSrc1, fileDes1Temp.decode('utf8'))
						#產生紀錄表
						createMHS(row, fileDes1Temp, p, i*6)

			except IOError as e:
				print u'無法產生 {0} 的紀錄表'.encode('big5').format(row[0])
				print u'進行中'.encode('big5')
				continue

			#建立拷貝zip
			try:
				shutil.copyfile(fileSrc2, fileDes2.decode('utf8'))
				#產生評估表
				createMHS(row, fileDes2)
			except IOError as e:
				print u'無法產生 {0} 的評估表'.encode('big5').format(row[0])
				print u'進行中'.encode('big5')
				continue

		i = i + 1

def createMHS(row, fileDes, p=None, recodeIndex=0):	
	#解壓縮
		zfr = zipfile.ZipFile(fileDes.decode('utf8'), 'r')
		data = zfr.read('word/document.xml')
		data = data.decode('utf8')
	#填入資料
		#病歷號碼
		data = re.sub(r'\b123456\b', row[0].decode('utf8'), data)

		#名字
		replaceLen = len(u'我的名字')
		nameLen = len(row[1].decode('utf8'))
		replaceName = row[1].decode('utf8') + u'　' * (replaceLen - nameLen)
		data = data.replace(u'我的名字', replaceName)

		#生日
		birth = row[2]
		if birth[0] == '0':
			birth = birth[1:]

		data = re.sub(r'\bbirth\b', birth.decode('utf8'), data)

		#初診日期
		firstDate = row[3].decode('utf8').split('/')
		data = re.sub(r'\bjkl\b', firstDate[0], data)
		data = re.sub(r'\bmn\b', firstDate[1], data)
		data = re.sub(r'\bop\b', firstDate[2], data)

		#男女
		if row[4] == u'男'.encode('utf8'):
			data = replaceOccurence(data, '<w:default w:val="1"/>', '', 1)
		elif row[4] == u'女'.encode('utf8'):
			data = replaceOccurence(data, '<w:default w:val="1"/>', '', 0)
		else:
			pass

		if p is not None:
			#紀錄
			records = row[6].split(',')

			for i in range(6):
				if i+recodeIndex < len(records):
					data = re.sub(r'\b00{0}\b'.format(i+1), records[i+recodeIndex], data)
					data = re.sub(r'\b01{0}\b'.format(i+1), str(p[0]), data)
					data = re.sub(r'\b02{0}\b'.format(i+1), str(p[1]), data)
					data = re.sub(r'\b03{0}\b'.format(i+1), str(p[2]), data)
					data = re.sub(r'\b04{0}\b'.format(i+1), str(p[3]), data)
					data = re.sub(r'\b05{0}\b'.format(i+1), str(p[4]), data)
					data = re.sub(r'\b06{0}\b'.format(i+1), str(p[5]), data)
					data = re.sub(r'\b07{0}\b'.format(i+1), str(p[6]), data)
					data = re.sub(r'\b08{0}\b'.format(i+1), str(p[7]), data)
					randomProgress(p)
				else:
					data = re.sub(r'\b00{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b01{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b02{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b03{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b04{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b05{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b06{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b07{0}\b'.format(i+1), '', data)
					data = re.sub(r'\b08{0}\b'.format(i+1), '', data)

		#刪除舊xml
		remove_from_zip(fileDes.decode('utf8'), 'word/document.xml')

		#寫入新xml
		zfw = zipfile.ZipFile(fileDes.decode('utf8'), 'a')
		zfw.writestr('word/document.xml', data.encode('utf8'))

		zfr.close()
		zfw.close()
	#改回docx
		filename = os.path.splitext(fileDes.decode('utf8'))[0]
		os.rename(fileDes.decode('utf8'), filename + ".docx")

print u'進行中'.encode('big5')
removeFileInFolder(outputPath)
parseInfo()
print u'結束'.encode('big5')
