import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
files_csv = [f for f in files if f[-3:] == 'txt']
print('number of files: '+str(len(files_csv)))
for work_item in range (0,len(files_csv)):
    IO = files_csv[work_item]
    #print(IO)
    # 輸入的檔案名稱，需先轉成 xls 檔，並下移一行
    #IO = 'SAA_table_all.csv'
    sheet = pd.read_csv(IO,header=None,sep="\s+")
    # 創建一個空白活頁簿物件
    wb = Workbook()
    # 選取正在工作中的表單
    ws = wb.active
    ws.cell(row = 1, column = 1, value = "m/z")
    ws.cell(row = 1, column = 2, value = "intensity")
    for i in range(0,len(sheet[0])):
        ws.cell(row = i+2, column = 1, value = sheet[0][i])
        ws.cell(row = i+2, column = 2, value = sheet[1][i])
    # 儲存成 create_sample.xls 檔案
    wb.save(IO[:-4] + '.xlsx')
    work_item = work_item + 1