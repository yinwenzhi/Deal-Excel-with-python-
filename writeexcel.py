from urllib.request import urlopen,Request
# from bs4 import BeautifulSoup
import re
import xlrd
import xlwt
from xlutils.copy import copy
import pandas as pd 
import time
import datetime

#xlrd只能读取excel
## 参考   https://www.cnblogs.com/nancyzhu/p/8401552.html
## https://www.cnblogs.com/jiangzhaowei/p/6179759.html

#xlwt只能写入excel

# #写excel
# #创建workbook（其实就是excel，后来保存一下就行）
# workbook = xlwt.Workbook(encoding = 'ascii')
# #创建表
# worksheet = workbook.add_sheet('My Worksheet')
# #往单元格内写入内容
# worksheet.write(0, 0, label = 'Row 0, Column 0 Value')
# workbook.save('Excel_Workbook.xls')


#xlutils.copy 可以修改exlel
# https://www.cnblogs.com/songdanqi0724/p/8145455.html 保留格式修改excel

#指定原始excel文件路径
excelfile='test.xls'
#pcklfile='C:\\Users\\VI\\Desktop\\result.txt'


class ExcelObject():
        def __init__(self,filepath):
                self.excelpath = filepath
                
                #读取excel
                #fiel_data = xlrd.open_workbook(excelfile)
                #formatting_info=True表示打开excel时并保存原有的格式
                self.file_data = xlrd.open_workbook(excelfile,formatting_info=True)

                #Sheet_data = data.sheets()[0]          #通过索引顺序获取工作簿
                #Sheet_data = data.sheet_by_index(0)    #通过索引顺序获取工作簿
                Sheet_data = self.file_data.sheet_by_name(u'Sheet1')    #通过名称获取工作簿 通过xlrd的sheet_by_index()获取的sheet没有write()方法

                #创建一个可以写入的副本
                #利用xlutils.copy函数，将xlrd.Book转为xlwt.Workbook，再用xlwt模块进行存储
                self.file_copy = copy(self.file_data)  

                self.Sheet_change = self.file_copy.get_sheet(0)#通过get_sheet()获取的sheet有write()方法

                #Sheet_change.write(0,0,'Row 0, column 0 Value')

                #使用pandas库传入该excel的数值仅仅是为了后续判断插入数据时应插入行是哪行
                self.original_data = pd.read_excel(excelfile,encoding='utf-8')
        
        #retain cell style
        def _getOutCell(self,writeSheet, colIndex, rowIndex):
                """ HACK: Extract the internal xlwt cell representation. """
                # self.setsheetchangebynum(sheetIndex)
                
                row = writeSheet._Worksheet__rows.get(rowIndex)
                # row = self.Sheet_change._Worksheet__rows.get(rowIndex)
                if not row: return None
                cell = row._Row__cells.get(colIndex)
                return cell

        #该函数中定义：对于没有任何修改的单元格，保持原有格式。
        def setOutCell(self,writeSheet, col, row, value):
                """ Change cell value without changing formatting. """
                

                # HACK to retain cell style.
                previousCell = self._getOutCell(writeSheet, col, row)
                # END HACK, PART I

                writeSheet.write(row, col, value)

                # HACK, PART II
                if previousCell:
                        newCell = self._getOutCell(writeSheet, col, row)
                        if newCell:
                                newCell.xf_idx = previousCell.xf_idx   
                         
        def append(self,sheetIndex,str_write):
                """从末尾写入excel数据"""
                row=len(self.original_data)+1
                str_write="添加写入"+str_write
                writeSheet = self.file_copy.get_sheet(sheetIndex)
                self.setOutCell(writeSheet,0,row,str_write)
                self.file_copy.save(self.excelpath)
                
                print('write: ',str_write)
                
        def write(self,sheetIndex,row,col,value):
                """根据sheetIndex、row、col写入value"""

                self.setOutCell(sheetIndex,col,row,value)
                # self.Sheet_change.write(row,col,value)
                self.file_copy.save(self.excelpath)

#设置表格样式 示例
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

#写Excel 示例
def write_excel():
    f = xlwt.Workbook()
    sheet1 = f.add_sheet('学生',cell_overwrite_ok=True)
    row0 = ["姓名","年龄","出生日期","爱好"]
    colum0 = ["张三","李四","恋习Python","小明","小红","无名"]
    #写第一行
    for i in range(0,len(row0)):
        #worksheet.write(1, 0, label = 'Formatted value', style) # Apply the Style to the Cell
        sheet1.write(0,i,row0[i],set_style('Times New Roman',220,True))
    #写第一列
    for i in range(0,len(colum0)):
        sheet1.write(i+1,0,colum0[i],set_style('Times New Roman',220,True))

    sheet1.write(1,3,'2006/12/12')
    sheet1.write_merge(6,6,1,3,'未知')#合并行单元格
    sheet1.write_merge(1,2,3,3,'打游戏')#合并列单元格
    sheet1.write_merge(4,5,3,3,'打篮球')

    f.save('test.xls')

excel  =  ExcelObject(excelfile)
excel.append(0,"b")

print('changed')


"""备忘录"""

# data = xlrd.open_workbook('excelfilepath') # 
# table = data.sheets()[0]
# table = data.sheet_by_index(0)
# table = data.sheet_by_name(u'Sheet1')
# table.row_values(i)
# table.col_values(i)
# nrows = table.nrows
# ncols = table.ncols
# sheetlist = table.sheets # A list of all sheets in the book.

# for i in range(nrows):
#   print(table.row_values(i))

# cell_A1 = table.cell(0,0).value
# cell_C4 = table.cell(2,3).value
"""


"""
# -*- coding: utf-8 -*-
# import xdrlib,sys
# import xlrd
# def open_excel(file='file.xls'):
# 	try:
# 		data = xlrd.open_workbook(file)
# 		return data
# 	except Exception,e:
		# print str(e)

"""根据索引获取Excel表格中的数据"""
#参数：file: Excel文件路径
#      colnameindex: 表头列名所在行的索引
#      by_index: 表的索引

# def excel_table_byindex(file='file.xls',colnameindex=0,by_index=0):
# 	data = open_excel(file)
# 	table = data.sheets()[by_index]
# 	nrows = table.nrows #行数
# 	ncols = table.ncols #列数
# 	colnames = table.row_values(colnameindex) #某一行数据
# 	list = []
# 	for rownum in range(1,nrows):
# 		row = table.row_values(rownum)#以列表格式输出
# 		if row:
# 			app = {}
# 			for i in range(len(colnames)):
# 				app[colnames[i]] = row[i]
# 			list.append(app)#向列表中插入字典类型的数据
# 	return list
 
# def main():
# 	tables = excel_table_byindex(file='test.xls')
# 	for row in tables:
# 		print row

# if __name__=="__main__":
# 	main()
""" """

"""通过名字索引"""
# # -*- coding: utf-8 -*-
# import xdrlib,sys
# import xlrd
# def open_excel(file='file.xls'):
# 	try:
# 		data = xlrd.open_workbook(file)
# 		return data
# 	except Exception,e:
# 		print str(e)

# def excel_table_byname(file='file.xls',colnameindex=0,by_name=u'Sheet1'):
# 	data = open_excel(file)
# 	table = data.sheet_by_name(by_name)
# 	nrows = table.nrows #行数
# 	colnames = table.row_values(colnameindex) #某一行数据
# 	list = []
# 	for rownum in range(1,nrows):
# 		row = table.row_values(rownum)
# 		if row:
# 			app = {}
# 			for i in range(len(colnames)):
# 				app[colnames[i]] = row[i]
# 			list.append(app)
# 	return list
 
# def main():
# 	tables = excel_table_byname(file='test.xls')
# 	for row in tables:
# 		print row

# if __name__=="__main__":
# 	main()

"""通过列表名索引"""
  # def getColumnIndex(table, columnName):
  #   columnIndex = None
  #     #print table
  #     for i in range(table.ncols):
  #       #print columnName
  #       #print table.cell_value(0, i)
  #       if(table.cell_value(0, i) == columnName):
  #         columnIndex = i
  #         break
  #   return columnIndex