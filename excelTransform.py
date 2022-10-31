import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
#IO = 'SAA_table_all.csv'
IO_source = 'train_org_filt.csv'
IO_target = 'train_table_filt.csv'
sheet_source = pd.read_csv(IO_source,header=None,sep=",")
sheet_target = pd.read_csv(IO_target,header=None, encoding = 'utf-8', sep=",", error_bad_lines=False)
# 創建一個空白活頁簿物件
wb = Workbook()
# 選取正在工作中的表單
ws = wb.active

# Store 24 Varient Name into varienNames list
title = ['SAA2a (RS-)', 'SAA2b (RS-)', 'SAA1g (RS-)', 'SAAU2 (RS-)', \
        'SAA1a (RS-)', 'SAAU1 (RS-)', 'SAA2a (R-)', 'SAA2b (R-)', \
        'SAA1b (RS-)', 'SAA1g (R-)', 'SAAU3 (RS-)', 'SAAU2 (R-)', \
        'SAA1a (R-)', 'SAAU1 (R-)', 'SAA1b (R-)', 'SAAU3 (R-)', \
        'SAA2a', 'SAA2b', 'SAA1g', 'SAAU2', 'SAA1a', 'SAAU1', 'SAA1b', 'SAAU3']
# Store 86 Bin Name into binNames list
for i in range(86):
    title.append('bin'+str(i+1))

print('file maximum rows: '+str(len(sheet_target.values)))
ws.cell(row = 1, column = 1, value = "test")
# assign title and initialize value
#for i in range(1,len(sheet.values)+1):
#    print('processing row:'+ str(i))
#    if ((i%3) == 1):
#        for j in range (0,4):
#            ws.cell(row = i+1, column = j+1, value = sheet.values[i-1][j])
find = 0
#for i in range(0, 6):
for i in range(0, len(sheet_target.values)):
    if (i%3 != 2):
        for j in range(0, len(sheet_target.values[i])):
            ws.cell(row = i+2, column = j+1, value = sheet_target.values[i][j])
    else:
        for j in range(0, len(sheet_target.values[i])):
            find = 0
            for k in range(0, len(title)):
                if sheet_target.values[i-2][j] == title[k]:
                    m = int(i/3) + 1
                    #print(sheet_source.values[m][k+1])
                    ws.cell(row = i+2, column = j+1, value = sheet_source.values[m][k+1])
                    find = 1
                    break
            if find == 0:
                ws.cell(row = i+2, column = j+1, value = sheet_target.values[i][j])

# 儲存成 create_sample.xls 檔案
os.remove(IO_target)
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv(IO_target, encoding='utf-8', sep=",",index=False,header=None)
os.remove('tmp.xls')