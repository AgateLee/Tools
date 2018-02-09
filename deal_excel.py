#!/usr/bin/python
# coding=utf-8

import sys 

reload(sys)
sys.setdefaultencoding('utf-8')

import xlwt;
import xlrd;
from xlutils.copy import copy;

def insert_date(file_name, date):
#	print file_name, date
	oldWb = xlrd.open_workbook(file_name, formatting_info=True);
	oldWbS = oldWb.sheet_by_index(0)
	newWb = copy(oldWb);
	newWs = newWb.get_sheet(0);
	nrows = oldWbS.nrows
	colNo = oldWbS.ncols
	cur = 1
	newWs.write(0, colNo, "time");
	while cur < nrows:
		newWs.write(cur, colNo, date);
		cur = cur + 1
	 
	newWb.save(file_name);

def GetFileList(FindPath,FlagStr=[]):  
	import os  
	FileList=[]  
	FileNames=os.listdir(FindPath)  
	if (len(FileNames)>0):  
	   for fn in FileNames:  
#		   fullfilename=os.path.join(FindPath,fn)  
		   FileList.append(fn)  
  
	if (len(FileList)>0):  
		FileList.sort()  
  
	return FileList  
	
def merge_all(file_path, dest):
	excels = GetFileList(file_path)

	first = True
	base = 0
	for item in excels:
		if not item[-3:] == 'xls':
			continue
		
		oldWb = xlrd.open_workbook(file_path + '/' + item, formatting_info=True);
		oldWbS = oldWb.sheet_by_index(0)
		
		print item, oldWbS.nrows
		
		if first:
			newWb = copy(oldWb);
			newWs = newWb.get_sheet(0);
			base = oldWbS.nrows
			first = False
		else:
			for rowIndex in range(1, oldWbS.nrows):
				for colIndex in range(oldWbS.ncols):
					newWs.write(base, colIndex, oldWbS.cell(rowIndex, colIndex).value)
				base = base + 1
	newWb.save(dest)
	
if __name__ == '__main__':
	
	lands = GetFileList('land')
	
	for item in lands:
		insert_date(r'land/' + item, item[:-4])
	merge_all('land', 'land.xls')
	
	lands = GetFileList('sales')
		
	for item in lands:
		insert_date(r'sales/' + item, item[:-4])
	merge_all('sales', 'sales.xls')
		
