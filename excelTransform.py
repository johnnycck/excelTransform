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
target.append(1)
FALSE = 0
TRUE = 1
ifFID = TRUE
ifACQU = TRUE
i = 0
path = path + '\\source'
wb = Workbook()
ws = wb.active
ws.cell(row = 1, column = 1, value = "test")
ws.cell(row = 2, column = 1, value = "missing file `fid`")
cur_row = 3
target_name = 'tmp'
for root, dir, file in os.walk(path):
    if os.path.basename(root) == '.git':
        dir[:] = []
    elif os.path.basename(root) == 'target':
        dir[:] = []
    elif os.path.basename(root) == '24 SAA + 50 bins':
        dir[:] = []
    else:
        if os.path.basename(root) != path:
            if ("ISD" in os.path.basename(root)):
                if(ifFID == FALSE):
                    ws.cell(row = cur_row, column = 1, value = target_name)
                    cur_row = cur_row+1
                if(ifACQU == FALSE):
                    ws.cell(row = cur_row, column = 1, value = target_name)
                    cur_row = cur_row+1
                target_name = os.path.basename(root)
                ifFID = FALSE
                ifACQU = FALSE
            for f in file:
                if ('fid' in f):
                    ifFID = TRUE
                if ('acqu' == f):
                    ifACQU = TRUE
        
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('missing dir.csv', encoding='utf-8',sep=',',index=False,header=None)
os.remove('tmp.xls')