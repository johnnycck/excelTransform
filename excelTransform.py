import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
dir_files = [f for f in files if os.path.isdir(f)]
dir_files = [f for f in dir_files if f[-4:] != '.git']
for i in range(len(dir_files)):
    new_string = old_string = dir_files[i]
    if 'IOC' in old_string:
        endIndex = old_string.find('IOC')
        tmp = old_string[:5]
        tmp = tmp + '(IOC)'
        tmp1 = old_string[5:endIndex]
        tmp1 = tmp1 + old_string[-1:]
        new_string = tmp + tmp1
    elif 'JZ' in old_string:
        endIndex = old_string.find('JZ')
        tmp = old_string[:5]
        tmp = tmp + '(JZ)'
        tmp1 = old_string[5:endIndex]
        tmp2 = old_string[endIndex+2:]
        new_string = tmp + tmp1 + tmp2
    os.rename(old_string, new_string)
    print(new_string)