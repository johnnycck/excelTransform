import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
files_dir = [f for f in files if f[-3:] != '.md']
files_dir = [f for f in files_dir if f[-3:] != '.py']
files_dir = [f for f in files_dir if f != '.git']
files_dir = [f for f in files_dir if f[-4:-3] != '.']
# 創建一個空白活頁簿物件
wb = Workbook()
# 選取正在工作中的表單
ws = wb.active
for work_item in range (0,len(files_dir)):
    IO = files_dir[work_item]
    print(IO)

    ws.cell(row = 1, column = 1, value = "test")
    print('---------- assign title ----------')
    # assign title
    ws.cell(row = 2, column = work_item+1, value = IO)
    
    print('---------- start extracting file name ----------')
    subPath = path + '/' + IO
    subDirFiles = os.listdir(subPath)
    cur_row = 3
    old_dir = []
    old_dir.append(1)
    cnt = 0
    for i in range(0,len(subDirFiles)):
        firstDashIndex = subDirFiles[i].index('_')
        tmp_front = subDirFiles[i][:firstDashIndex]
        if (subDirFiles[i][-3:] == 'csv'):
            tmp_back = subDirFiles[i][firstDashIndex:-4]
        else:
            tmp_back = subDirFiles[i][-2:]
        tmp = tmp_front + tmp_back
        print(tmp)
        '''
        hasDuplicate = 0
        for j in range(0,cnt-1):
            if(tmp == old_dir[j]):
                hasDuplicate = 1
                break
        if(hasDuplicate):
        '''
        #ws.cell(row = cur_row, column = work_item+1, value = subDirFiles[i])
        ws.cell(row = cur_row, column = work_item+1, value = tmp)
        cur_row = cur_row+1
        #else:
        #    old_dir[cnt] = tmp
        #    old_dir.append(1)
        #    cnt = cnt + 1


    print('dir number '+str(work_item+1) +' is finished')
    work_item = work_item + 1
# 儲存成 create_sample.xls 檔案
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('batch統計表.csv', encoding='utf-8',sep=',',index=False,header=None)
os.remove('tmp.xls')