import os
import xlwt
from xlutils import copy
import xlrd
import os
import win32com
from win32com.client import constants as c 
import pandas as pd

from xlutils.copy import copy

from xlwt import Style
a = os.getcwd() #获取当前目录
print (a) #打印当前目录
os.chdir(r'C:\Users\PC\Desktop') #定位到新的目录，请根据你自己文件的位置做相应的修改
a = os.getcwd() #获取定位之后的目录
print(a) #打印定位之后的目录
#读取目标txt文件里的内容，并且打印出来显示
with open(r'C:\Users\PC\Desktop\Q2Kv2_12b1\demo.txt','r') as raw:
	for line in raw:
		print (line)

#创建一个workbook对象，相当于创建一个Excel文件
# book = xlwt.Workbook(encoding='utf-8',style_compression=0)

# 创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格。
# sheet = book.add_sheet('Output', cell_overwrite_ok=True)
#其中的Output是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False


rb = xlrd.open_workbook(r"C:/Users/PC\Desktop/Output.xls", formatting_info=True)
wb = copy(rb)
sheet = wb.get_sheet(0)

# 向表中添加数据标题
sheet.write(0, 0, 'Distance')  # 其中的'0-行, 0-列'指定表中的单元，'X'是向该单元写入的内容
sheet.write(0, 1, 'Mean')
sheet.write(1, 0, 'x(km)')  # 其中的'0-行, 0-列'指定表中的单元，'X'是向该单元写入的内容
sheet.write(1, 1, 'Temp-data')

#对文本内容进行多次切片得到想要的部分
n=2
with open(r'C:\Users\PC\Desktop\Q2Kv2_12b1\demo.txt',) as fd:
	for text in fd.readlines():
		x=text.split()[0]
		y=text.split()[1]
		print (x)
		print (y)
		sheet.write(n,0,x)#往表格里写入X坐标 
		sheet.write(n,1,y)#往表格里写入Y坐标
		n = n+1
# 最后，将以上操作保存到指定的Excel文件中
wb.save('Output.xls')  


