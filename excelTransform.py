import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook

path = os.getcwd()
files = os.listdir(path)
files_txt = [f for f in files if f[-3:] == 'txt']
print('number of files: '+str(len(files_txt)))

criterion = ['GC(GC)', 'GC(non-GC)', 'GC(KN)', 'non-GC(GC)', 'non-GC(non-GC)', 'non-GC(KN)', 'KN(GC)', 'KN(non-GC)', 'KN(KN)']
criterion_num = 9

# create an empty work sheet
wb = Workbook()
# select active sheet
ws = wb.active
# Initial tmp title, which will be deleted when creating real output csv
ws.cell(row = 1, column = 1, value = "test")
# Initial title
ws.cell(row = 2, column = 1, value = "檔名")
ws.cell(row = 2, column = 2, value = "Threshold")

for col in range(0,criterion_num):
    ws.cell(row = 2, column = col+3, value = criterion[col])
cur_row = 3

for work_item in range (0,len(files_txt)):
    IO = files_txt[work_item]
    sheet = open(IO)
    # initial txt file info
    file_row_num = 0
    file_text = []
    for line in sheet:
        file_text.append(line)
        file_row_num += 1
    # dec. file_row_num for empty line
    for i in range(file_row_num-1, 0, -1):
        if('|' in file_text[i]):
            file_row_num = i
            break
    # retrieve 'GC' info
    row1 = file_text[file_row_num-4]
    # retrieve 'non-GC' info
    row2 = file_text[file_row_num-2]
    # retrieve 'KN' info
    row3 = file_text[file_row_num]
    output = [[0]*3 for i in range(0,9)]
    # 19, 29, 39, 49 means '|' index of the row
    # need to get digit between '|' and '|'

    # retrieve 'GC' info
    output[0] = [f for f in row1[19:29] if f.isdigit()]
    output[1] = [f for f in row1[29:39] if f.isdigit()]
    output[2] = [f for f in row1[39:49] if f.isdigit()]
    # retrieve 'non-GC' info
    output[3] = [f for f in row2[19:29] if f.isdigit()]
    output[4] = [f for f in row2[29:39] if f.isdigit()]
    output[5] = [f for f in row2[39:49] if f.isdigit()]
    # retrieve 'KN' info
    output[6] = [f for f in row3[19:29] if f.isdigit()]
    output[7] = [f for f in row3[29:39] if f.isdigit()]
    output[8] = [f for f in row3[39:49] if f.isdigit()]
    
    # merge digit into integer
    for i in range(0,9):
        if(len(output[i]) == 1):
            output[i] = output[i][0]
        if(len(output[i]) == 2):
            output[i] = output[i][0] + output[i][1]
        if(len(output[i]) == 3):
            output[i] = output[i][0] + output[i][1] + output[i][2]
    
    # write output
    ws.cell(row = cur_row, column = 1, value = IO[:-15])
    ws.cell(row = cur_row, column = 2, value = IO[-15:-4])
    ws.cell(row = cur_row, column = i+3, value = output[i])
    for i in range(0,9):
        ws.cell(row = cur_row, column = i+3, value = output[i])
    
    cur_row += 1
    work_item = work_item + 1
    print('finish file:'+str(work_item))

# store as tmp xls file first
wb.save('tmp.xls')
# transform xls to csv
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('output.csv', encoding='ANSI',sep=',',index=False,header=None)
os.remove('tmp.xls')