#coding=utf-8
 
import os
import time
import requests
from xlrd import open_workbook
import xlwt 
from xlutils.copy import copy



date=time.strftime("%M%S%H%d%m%y")
date_for_file=time.strftime("%Y%m%d%H")

# 儲存路徑資料夾/不存在則新增一資料夾
path=r"EurexMargin\\"+time.strftime("%y%m")+"\\"

if not os.path.isdir(r"EurexMargin\\"):
	os.mkdir(r"EurexMargin\\")

if not os.path.isdir(path):
	os.mkdir(path)
# 下載資料
file_name=("Eurex Margin_"+time.strftime("%Y")+time.strftime("%m")+time.strftime("%d")+".xls")
dls = "https://www.eurexchange.com/resource/blob/337412/624bdec070f9df4432fc2caa2558668e/MarginParametersEstimationCircular-data.xls"
resp = requests.get(dls)

output = open(path+file_name, 'wb')
output.write(resp.content)
output.close()

# 寫入excel檔
rb = open_workbook(path+"Eurex Margin_"+time.strftime("%Y")+time.strftime("%m")+time.strftime("%d")+".xls")
wb = copy(rb)
sheet = rb.sheet_by_index(0)
s = wb.get_sheet(0)
col8 = sheet.col_values(8)
col2 = sheet.col_values(3)

# 依儲存格寫入所需公式

s.write(0,11,u'參考')
s.write(0,12,u'原始&維持')
s.write(0,13,u'當沖')
num1 = int(col2.index("FDAX")+1)
s.write(1,10,'DAX')
s.write(1,11, xlwt.Formula('I'+str(num1)))
s.write(1,12, xlwt.Formula('ROUND(L2*1.11,0)'))
s.write(1,13, xlwt.Formula('ROUND(M2/2,2)'))

num2 = int(col2.index("FDXM")+1)
s.write(2,10,'FDXM')
s.write(2,11, xlwt.Formula('I'+str(num2)))
s.write(2,12, xlwt.Formula('ROUND(L3*1.11,0)'))
s.write(2,13, xlwt.Formula('ROUND(M3/2,2)'))

num3 = int(col2.index("FESB")+1)
s.write(3,10,'FESB')
s.write(3,11, xlwt.Formula('I'+str(num3)))
s.write(3,12, xlwt.Formula('ROUND(L4*1.11,0)'))
s.write(3,13, xlwt.Formula('ROUND(M4/2,2)'))

num4 = int(col2.index("FESX")+1)
s.write(4,10,'FESX')
s.write(4,11, xlwt.Formula('I'+str(num4)))
s.write(4,12, xlwt.Formula('ROUND(L5*1.11,0)'))
s.write(4,13, xlwt.Formula('ROUND(M5/2,2)'))

num5 = int(col2.index("FGBL")+1)
s.write(5,10,'FGBL')
s.write(5,11, xlwt.Formula('I'+str(num5)))
s.write(5,12, xlwt.Formula('ROUND(L6*1.11,0)'))
s.write(5,13, xlwt.Formula('ROUND(M6/2,2)'))

num6 = int(col2.index("FGBM")+1)
s.write(6,10,'FGBM')
s.write(6,11, xlwt.Formula('I'+str(num6)))
s.write(6,12, xlwt.Formula('ROUND(L7*1.11,0)'))
s.write(6,13, xlwt.Formula('ROUND(M7/2,2)'))

num7 = int(col2.index("FGBS")+1)
s.write(7,10,'FGBS')
s.write(7,11, xlwt.Formula('I'+str(num7)))
s.write(7,12, xlwt.Formula('ROUND(L8*1.11,0)'))
s.write(7,13, xlwt.Formula('ROUND(M8/2,2)'))

num8 = int(col2.index("FGBX")+1)
s.write(8,10,'FGBX')
s.write(8,11, xlwt.Formula('I'+str(num8)))
s.write(8,12, xlwt.Formula('ROUND(L9*1.11,0)'))
s.write(8,13, xlwt.Formula('ROUND(M9/2,2)'))

num9 = int(col2.index("FBTP")+1)
s.write(9,10,'FBTP')
s.write(9,11, xlwt.Formula('I'+str(num9)))
s.write(9,12, xlwt.Formula('ROUND(L10*1.11,0)'))
s.write(9,13, xlwt.Formula('ROUND(M10/2,2)'))

num10 = int(col2.index("FOAT")+1)
s.write(10,10,'FOAT')
s.write(10,11, xlwt.Formula('I'+str(num10)))
s.write(10,12, xlwt.Formula('ROUND(L11*1.11,0)'))
s.write(10,13, xlwt.Formula('ROUND(M11/2,2)'))

wb.save(path+"Eurex Margin_"+time.strftime("%Y")+time.strftime("%m")+time.strftime("%d")+".xls")
