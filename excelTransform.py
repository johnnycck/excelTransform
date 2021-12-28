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
n_rows, n_cols = input.shape
for i in range(0,n_rows):
    target[i] = input[i][1]
    if i != n_rows-1:
        target.append(1)
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
                    if (target[i] in f):
                        #print(f)
                        print(root)
                        s_path = root + '\\' + f
                        t_file = t_path + '\\' + f
                        shutil.copy(s_path, t_file)