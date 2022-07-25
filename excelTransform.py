import pandas as pd
import math
import os
import numpy as np
import openpyxl
import sys
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
files_csv = [f for f in files if f[-3:] == 'csv']
range_low = 11300
range_high = 11800
print('number of files: '+str(len(files_csv)))
# 批次處理同資料夾內的所有 csv 檔
for work_item in range (0,len(files_csv)):
    wb = Workbook()
    ws = wb.active
    IO = files_csv[work_item]
    sheet = pd.read_csv(IO,header=None,sep=",")
    # 填入前三行的資訊
    ws.cell(row = 2, column = 1, value = sheet.values[0][0])
    ws.cell(row = 3, column = 1, value = sheet.values[0][0])
    ws.cell(row = 2, column = 2, value = sheet.values[0][1])
    ws.cell(row = 3, column = 2, value = sheet.values[0][1])
    ws.cell(row = 2, column = 3, value = sheet.values[0][2])
    ws.cell(row = 3, column = 3, value = sheet.values[0][2])
    print('---- start tracking start index ----')
    start = 3
    end = -1
    # 先從後面 scan，因後面開始算會比較接近 11300，並以一次跳10個的方式來省時間
    for j in range(len(sheet.values[1])-1,3,-10):
        if (sheet.values[0][j] < range_high and end == -1):
            end = j
        if (sheet.values[0][j] < range_low):
            start = j - 1
            break
    print('---- start storing data within desired range ----')
    total = (end - start)/100
    col=4
    for j in range(start,len(sheet.values[1])):
        if (sheet.values[0][j] >= range_low and sheet.values[0][j] <= range_high):
            ws.cell(row = 2, column = col, value = sheet.values[0][j])
            ws.cell(row = 3, column = col, value = sheet.values[1][j])
            print('---- progress: ' + str(int((j - start)/total)) + '% ----', end = "\r")
            col = col+1
        if (sheet.values[0][j] > range_high):
            break
    # 存檔並刪除舊的 CSV
    print('\nfile number '+str(work_item+1) +' has finished')
    wb.save('tmp.xls')
    data_xls = pd.read_excel('tmp.xls',index_col=None)
    os.rename(IO, 'tmp')
    data_xls.to_csv(IO, encoding='utf-8',sep=',',index=False,header=None)
    os.remove('tmp.xls')
    os.remove('tmp')
    work_item = work_item + 1