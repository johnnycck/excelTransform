import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook
import shutil

path = os.getcwd()
target = []
target_name = []
target.append(1)
target_name.append(1)
wb = Workbook()
ws = wb.active
ws.cell(row = 1, column = 1, value = "test")
ws.cell(row = 2, column = 1, value = "IOC batch name")
ws.cell(row = 2, column = 2, value = "IOC batch number")
ws.cell(row = 2, column = 3, value = "JZ batch name")
ws.cell(row = 2, column = 4, value = "JZ batch number")
JZ_row_num = 3
IOC_row_num = 3
IOC = 0
JZ = 0
d_JZ = 0
path = path + '\\source'
for root, dir, file in os.walk(path):
    IOC = 0
    JZ = 0
    if os.path.basename(root) == '.git':
        dir[:] = []
    elif os.path.basename(root) == 'target':
        dir[:] = []
    elif ('B' in root):
        b_num = root.index('B')
        if ('IOC' in root):
            JZ = 0
        if ('JZ' in root):
            JZ = 1
        #print(root)
        for d in dir:
            b_name_end = d.index('_')
            d_JZ = 0
            if ('JZ' in d):
                d_JZ = 1
            if (JZ == 1 or d_JZ == 1):
                ws.cell(row = JZ_row_num, column = 3, value = d[:b_name_end])
                ws.cell(row = JZ_row_num, column = 4, value = root[b_num:b_num+4])
                JZ_row_num += 1
            else:
                ws.cell(row = IOC_row_num, column = 1, value = d[:b_name_end])
                ws.cell(row = IOC_row_num, column = 2, value = root[b_num:b_num+4])
                IOC_row_num += 1
            #print(root[b_num:b_num+4])
            #print(d[:b_name_end])
        dir[:] = []
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('batch name and number.csv', encoding='utf-8',sep=',',index=False,header=None)
os.remove('tmp.xls')