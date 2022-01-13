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
for i in range(0,n_rows):
    target[i] = input[i][1]
    target_name[i] = 0
    if i != n_rows-1:
        target.append(1)
        target_name.append(1)
path = path + '\\source'
already_copy = []
n_dir = 0
for root, dir, file in os.walk(path):
    if os.path.basename(root) == '.git':
        dir[:] = []
    elif os.path.basename(root) == 'target':
        dir[:] = []
    elif os.path.basename(root) == '__MACOSX':
        dir[:] = []
    else:
        if os.path.basename(root) != path:
            for f in dir:
                for i in range(0, n_rows):
                    if (target[i] in f):
                        s_path = root + '\\' + f
                        t_file = t_path + '\\' + f
                        print(f)
                        al_copy = 0
                        for iter_dir in range(0, n_dir):
                            if (already_copy[iter_dir] == f):
                                al_copy = 1
                                break
                        if(al_copy == 0):
                            shutil.copytree(s_path, t_file)
                            already_copy.append(1)
                            already_copy[n_dir] = f
                            n_dir = n_dir+1
                        target_name[i] = 1
wb = Workbook()
ws = wb.active
ws.cell(row = 1, column = 1, value = "test")
ws.cell(row = 2, column = 1, value = "missing dir name")
cur_row = 3
for i in range(0, n_rows):
    if(target_name[i] == 0):
        ws.cell(row = cur_row, column = 1, value = target[i]) 
        cur_row = cur_row+1
        
wb.save('tmp.xls')
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('missing dir.csv', encoding='utf-8',sep=',',index=False,header=None)
os.remove('tmp.xls')