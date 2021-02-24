import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
files_csv = [f for f in files if f[-3:] == 'csv']
print('number of files: '+str(len(files_csv)))
for work_item in range (0,len(files_csv)):
    IO = files_csv[work_item]
    #print(IO)
    # 輸入的檔案名稱，需先轉成 xls 檔，並下移一行
    #IO = 'SAA_table_all.csv'
    sheet = pd.read_csv(IO,header=None,sep=",")
    # 創建一個空白活頁簿物件
    wb = Workbook()
    # 選取正在工作中的表單
    ws = wb.active
    # title
    titles = ['PID', 'CANCER', 'CANCER1', 'LABEL', 'MZ', 'INTENSITY']
    ISD_1 = []
    ISD_2 = []
    MAD = []
    print('file maximum rows: '+str(len(sheet.values)))
    ws.cell(row = 1, column = 1, value = "test")
    print('---------- assign title ----------')
    # assign title
    for i in range(0,6):
        ws.cell(row = 2, column = i+1, value = titles[i])
    
    print('---------- start transposing content ----------')
    cur_row = 3
    for i in range(0,len(sheet.values),3):
        q_num = 0 # '?' number
        print('processing row: '+ str(i))
        for j in range(4,114):
            ws.cell(row = cur_row, column = 1, value = sheet.values[i][1]) 
            ws.cell(row = cur_row, column = 4, value = sheet.values[i][j]) 
            ws.cell(row = cur_row, column = 5, value = sheet.values[i+1][j])
            ws.cell(row = cur_row, column = 6, value = sheet.values[i+2][j])
            cur_row = cur_row+1
    print('file number '+str(work_item+1) +' is finished')
    # 儲存成 create_sample.xls 檔案
    wb.save('tmp.xls')
    data_xls = pd.read_excel('tmp.xls',index_col=None)
    if('.csv' in files_csv[work_item]):
        files_csv[work_item] = files_csv[work_item][:-4]
    data_xls.to_csv('分析_'+files_csv[work_item]+'.csv', encoding='utf-8',sep=',',index=False,header=None)
    os.remove('tmp.xls')
    work_item = work_item + 1