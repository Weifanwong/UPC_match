#encoding=utf-8
import xlrd
from xlwt import *
from openpyxl import Workbook as wb
import os
import re
# headlen = len(table_head)
# for i in range(headlen):
# 	xlsheet.cell(row=1,column=i+1).value = table_head[i] #写表头

n1 = '!@#dsad 1234'
#n1 = n1.strip(' ')
n1 = re.sub('[^0-9]','',n1)
print(n1.isdigit())

