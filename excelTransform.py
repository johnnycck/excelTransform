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
    MZnum = 0;
    print('---------- check number of m/z between 5000~12450 ----------')
    for i in range(0,len(sheet[0])):
        if(sheet[0][i] >= 5000 and sheet[0][i] <=12450):
            MZnum = MZnum + 1
    print('MZnum: '+str(MZnum))
    if(MZnum > 16384):
        MZnum = 0;
        print('---------- check number of m/z between 11300~12450 ----------')
        for i in range(0,len(sheet[0])):
            if(sheet[0][i] >= 11300 and sheet[0][i] <=12450):
                MZnum = MZnum + 1
        if(MZnum > 16384):
            print('MZnum: '+str(MZnum))
            print('the number of m/z between 11300~12450 for file: '+ files_csv[work_item] +'exceed excel limitation')
        else:
            print('MZnum: '+str(MZnum))
            # 創建一個空白活頁簿物件
            wb = Workbook()
            # 選取正在工作中的表單
            ws = wb.active
            # Initial title
            batch_name = IO.split('_')
            titles = ['All Patient Data', batch_name[0], IO[:-4]]

            ws.cell(row = 1, column = 1, value = "test")
            print('---------- assign title ----------')
            # assign title
            for i in range(0,3):
                ws.cell(row = 2, column = i+1, value = titles[i])
                ws.cell(row = 3, column = i+1, value = titles[i])
            
            print('---------- start transposing content ----------')
            col_num = 4
            for i in range(3,len(sheet[0])+2):
                if(sheet[0][i-3] >= 11300 and sheet[0][i-3] <=12450):
                    ws.cell(row = 2, column = col_num, value = round(sheet[0][i-3],1))
                    ws.cell(row = 3, column = col_num, value = sheet[1][i-3])
                    col_num = col_num + 1
            print('file number '+str(work_item+1) +' is finished')
            
            # 儲存成 create_sample.xls 檔案
            wb.save('tmp.xls')
    
            data_xls = pd.read_excel('tmp.xls',index_col=None)
            if('.txt' in files_csv[work_item]):
                files_csv[work_item] = files_csv[work_item][:-4]
            data_xls.to_csv(files_csv[work_item]+'.csv', encoding='utf-8',sep=',',index=False,header=None)
            os.remove('tmp.xls')
    else:
        print('MZnum: '+Mznum)
        # 創建一個空白活頁簿物件
        wb = Workbook()
        # 選取正在工作中的表單
        ws = wb.active
        # Initial title
        batch_name = IO.split('_')
        titles = ['All Patient Data', batch_name[0], IO[:-4]]

        ws.cell(row = 1, column = 1, value = "test")
        print('---------- assign title ----------')
        # assign title
        for i in range(0,3):
            ws.cell(row = 2, column = i+1, value = titles[i])
            ws.cell(row = 3, column = i+1, value = titles[i])
            
        print('---------- start transposing content ----------')
        col_num = 4
        for i in range(3,len(sheet[0])+2):
            if(sheet[0][i-3] >= 5000 and sheet[0][i-3] <=12450):
                ws.cell(row = 2, column = col_num, value = round(sheet[0][i-3],1))
                ws.cell(row = 3, column = col_num, value = sheet[1][i-3])
                col_num = col_num+1
        print('file number '+str(work_item+1) +' is finished')
            
        # 儲存成 create_sample.xls 檔案
        wb.save('tmp.xls')
    
        data_xls = pd.read_excel('tmp.xls',index_col=None)
        if('.txt' in files_csv[work_item]):
            files_csv[work_item] = files_csv[work_item][:-4]
        data_xls.to_csv(files_csv[work_item]+'.csv', encoding='utf-8',sep=',',index=False,header=None)
        os.remove('tmp.xls')
    work_item = work_item + 1