import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
files_csv = [f for f in files if f[-10:] == 'result.csv']

# 創建一個空白活頁簿物件
wb = Workbook()
# 選取正在工作中的表單
ws = wb.active

print('number of files: '+str(len(files_csv)))
print(files_csv)
titles = ['Sensitivity', 'Specifically', 'Accuracy']
print('---------- assign title ----------')
# assign title
for i in range(0,3):
    ws.cell(row = 2, column = i+2, value = titles[i])
cancers = ['GC', 'CR', 'LC', 'EC', 'GEO', 'OC']
ws.cell(row = 1, column = 1, value = "test")
for work_item in range (0,len(files_csv)):
    IO = files_csv[work_item]
    #print(IO)
    # 輸入的檔案名稱，需先轉成 xls 檔，並下移一行
    #IO = 'SAA_table_all.csv'
    sheet = pd.read_csv(IO,header=None,sep=",")
    TP = 0
    FP = 0
    FN = 0
    TN = 0
    
    print('file maximum rows: '+str(len(sheet.values)))
    ws.cell(row = work_item+3, column = 1, value = files_csv[work_item][:-4])
    print('---------- start processing ----------')
    cur_row = 3
    #print(sheet.values[1][0])
    ifCancer = 0
    for i in range(1,len(sheet.values)):
        for j in range(0,6):
            if(cancers[j] in sheet.values[i][0]):
                ifCancer = 1
                break
            else:
                ifCancer = 0

        if(ifCancer==1 and int(sheet.values[i][1])==1):
            TP = TP + 1
        elif(ifCancer==1 and int(sheet.values[i][1])==0):
            FN = FN + 1
        elif(ifCancer==0 and int(sheet.values[i][1])==1):
            FP = FP + 1
        elif(ifCancer==0 and int(sheet.values[i][1])==0):
            TN = TN + 1
    # Sensitivity
    ws.cell(row = work_item+3, column = 2, value = round((TP/(TP+FN))*100,2))
    # Specifically
    ws.cell(row = work_item+3, column = 3, value = round((TN/(FP+TN))*100,2))
    # Accuracy
    ws.cell(row = work_item+3, column = 4, value = round(((TP+TN)/(len(sheet.values)-1))*100,2))
    print('file number '+str(work_item+1) +' is finished')
    work_item = work_item + 1
    # 儲存成 create_sample.xls 檔案
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('分析表.csv', encoding='utf-8',sep=',',index=False,header=None)
os.remove('tmp.xls')