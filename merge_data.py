#!/usr/bin/python
# coding=utf-8

import sys 
import os
import re

reload(sys)
sys.setdefaultencoding('utf-8')

import xlwt;
import xlrd;
from xlutils.copy import copy;

infos = {
	u'销售套数(套)': 5,
	u'销售面积(㎡)': 6,
	u'销售均价(元/㎡)': 7,
	u'销售金额(万元)': 8,
	u'上市套数(套)': 9,
	u'上市面积(㎡)': 10,
}


def GetFileList(FindPath, AllFiles=True,FlagStr=[]):    
	print 'READ PATH : %s' % FindPath
	FileList=[]  
	FileNames=os.listdir(FindPath)  
	if (len(FileNames)>0):  
	    for fn in FileNames:    
			if AllFiles | os.path.isdir(FindPath + '/' +fn):
				FileList.append(fn)
			else:
				print 'JUMP %s' % fn
  
	if (len(FileList)>0):  
		FileList.sort()  
  
	return FileList  
	
	
def AnalyseName(file_name):
	m = re.match(ur'^([a-zA-Z0-9\.\u4e00-\u9fa5]*[a-zA-Z\u4e00-\u9fa5]+)([0-9]*)([ -])?([\u4e00-\u9fa5\_]*)?([ -])?((\d{2,4})(-\d{2,4})?)?( )?(.xls)$', file_name)
	
	return m.group(1), m.group(2), m.group(4), m.group(6)
	
	
def checkSheet(file_name, sheet, base, offset, cat, company, number):
	oldWb = xlrd.open_workbook(file_name, formatting_info=True);
	oldWbS = oldWb.sheet_by_index(0)
	
	if infos.has_key(oldWbS.cell(0, 1).value):
		print 'No City: ' + file_name
		return 0
	elif oldWbS.cell(2, 0).value != u'总计':
		print 'No Zongji: ' + file_name
		st_x = 2
	else:
		st_x = 3
		
	st_y = 1
	old_base = base;
	mapping = {}
#	print '###',offset
	for sed in range (0, offset):
		try:
			mapping[sed] = infos[oldWbS.cell(1, sed + 1).value]
		except Exception as e:
			print oldWbS.cell(1, sed + 1).value

	while st_y < oldWbS.ncols:
		city = oldWbS.cell(0, st_y).value
#		print 'sovle', city
		for i in range(st_x, oldWbS.nrows):
#			print 'start from %d -> result %d' % (i, base)	
			sheet.write(base, 0, cat)
			sheet.write(base, 1, number)
			sheet.write(base, 2, company)
			sheet.write(base, 3, oldWbS.cell(i, 0).value)
			sheet.write(base, 4, city)
			for j in range(0, offset):
				sheet.write(base, mapping[j], oldWbS.cell(i, st_y + j).value)
			base = base + 1;
		st_y = st_y + offset
			
	return base - old_base
	
def printSheet(file_name):
	oldWb = xlrd.open_workbook(file_name, formatting_info=True);
	oldWbS = oldWb.sheet_by_index(0)
		
	for rowIndex in range(0, oldWbS.nrows):
		for colIndex in range(oldWbS.ncols):
			print '[' + oldWbS.cell(rowIndex, colIndex).value + ']',
		print
	
	
if __name__ == '__main__':
	
	workbook = xlwt.Workbook() 
	
	sheets = []
	for i in range(0,10):
		sheet = workbook.add_sheet('list' + str(i), cell_overwrite_ok=True) 
		sheet.write(0, 0, u'分类')
		sheet.write(0, 1, u'代码')
		sheet.write(0, 2, u'企业名称')
		sheet.write(0, 3, u'月度')
		sheet.write(0, 4, u'城市')
		sheet.write(0, 5, u'销售套数(套)')
		sheet.write(0, 6, u'销售面积(㎡)')
		sheet.write(0, 7, u'销售均价(元/㎡)')
		sheet.write(0, 8, u'销售金额(万元)')
		sheet.write(0, 9, u'上市套数(套)')
		sheet.write(0, 10, u'上市面积(㎡)')
		sheets.append(sheet)
		
	new_line = 1
	add_lines = 0
	idx = 0
	tups = {}
	
	folders = GetFileList('data', False)
	
	for item in folders:
#	for item in ['测试']:
		data_list = GetFileList('data/' + item)
		
		for file_name in data_list:
			if file_name.endswith('.xls'):
				print 'Process %s' % file_name
				
				company, number, content, year = AnalyseName(file_name.decode('utf-8'))

				tup = (company, number, year) 
				
				if tups.has_key(tup):
					base = tups[tup]
				else:
					base = new_line
					
				if content == '':
					add_lines = checkSheet('data/' + item + '/' + file_name, sheets[idx], base, 6, 
						item.decode('utf-8'), company, number)
				else:
					add_lines = checkSheet('data/' + item + '/' + file_name, sheets[idx], base, len(content.split('_')), 
						item.decode('utf-8'), company, number)
						
#				print '###',add_lines
				if tups.has_key(tup):
					continue
				else:
					tups[tup] = base
					new_line = base + add_lines if isinstance(add_lines, int) else 0
				
				if new_line > 50000:
					idx = idx + 1
					new_line = 1
					workbook.save('data/result.xls')
#				break
			else:
				print 'JUMP %s' % file_name
#		break
				
	workbook.save('data/result.xls')
		