import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook
import shutil

path = os.getcwd()
t_path = path + '\\target'
files = os.listdir(path)
input_xlsx = pd.read_excel('input.xlsx')
input = input_xlsx.values
target = []
target_name = []
target.append(1)
target_name.append(1)
n_rows, n_cols = input.shape
wb = Workbook()
ws = wb.active
ws.cell(row = 1, column = 1, value = "test")
ws.cell(row = 2, column = 1, value = "csv name")
ws.cell(row = 2, column = 2, value = "batch number")
row_num = 3;
for i in range(0,n_rows):
    target[i] = input[i][1]
    target_name[i] = 0
    if i != n_rows-1:
        target.append(1)
        target_name.append(1)
path = path + '\\source'
for root, dir, file in os.walk(path):
    if os.path.basename(root) == '.git':
        dir[:] = []
    elif os.path.basename(root) == 'target':
        dir[:] = []
    else:
        if os.path.basename(root) != path:
            for f in file:
                for i in range(0, n_rows):
                    if (target[i] in f and target_name[i] == 0):
                        #print(f)
                        BatchStart = root.index('B')
                        print(root[BatchStart:BatchStart+4])
                        print(f[:-4])
                        s_path = root + '\\' + f
                        t_file = t_path + '\\' + f
                        shutil.copy(s_path, t_file)
                        ws.cell(row = row_num, column = 1, value = f[:-4])
                        ws.cell(row = row_num, column = 2, value = root[BatchStart:BatchStart+4])
                        row_num +=1 ;
                        target_name[i] = 1
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('csv and batch number.csv', encoding='utf-8',sep=',',index=False,header=None)
os.remove('tmp.xls')

wb = Workbook()
ws = wb.active
ws.cell(row = 1, column = 1, value = "test")
ws.cell(row = 2, column = 1, value = "csv name")
cur_row = 3
for i in range(0, n_rows):
    if(target_name[i] == 0):
        ws.cell(row = cur_row, column = 1, value = target[i]) 
        cur_row = cur_row+1
        
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('missing.csv', encoding='utf-8',sep=',',index=False,header=None)
os.remove('tmp.xls')