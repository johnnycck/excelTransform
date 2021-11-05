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

criterion = ['TP', 'TN', 'TP+TN', 'PPV', 'NPV', 'Accuracy', 'Sensitivity', 'Specificity']
criterion_num = 8
print('number of criterion: '+str(criterion_num))

# initial user_command criterion list
criterion_user_input = []
criterion_user_input = [[0]*2 for i in range(8)]
for i in range(0, criterion_num):
    criterion_user_input[i][0] = criterion[i]
    criterion_user_input[i][1] = -1
# type user command from stdin
print('Please type your criterion, type **-1** for not specifying this criterion:')
for i in range(0, criterion_num):
    criterion_user_input[i][1] = input(criterion[i] + ': ')

# only deal with real criterion(not -1)
target_criterion = []
target_criterion_num = 0
for i in range(0, criterion_num):
    if(not('-1' in criterion_user_input[i][1])):
        target_criterion.append(i)
        target_criterion_num += 1
if(target_criterion_num == 0):
    quit()
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
    
    # classify text into items
    file_textbox = [[0]*9 for i in range(file_row_num-3)]
    row = 0
    for i in range(4, file_row_num+1):
        lastk = k = 1; # skip first '|'
        for j in range(0, 9):
            while(1):
                if(file_text[i][k] == ' '):
                    k += 1
                else:
                    break
            lastk = k
            while(1):
                if(file_text[i][k] != ' '):
                    k += 1
                else:
                    break
            file_textbox[row][j] = file_text[i][lastk:k]
            while(1):
                if(file_text[i][k] == ' '):
                    k += 1
                else:
                    break
            k += 1
        row += 1
    file_row_num -= 3
    # find match criterion and store to csv file
    file_row_pass = True
    for i in range(0, file_row_num):
        file_row_pass = True
        for j in range(0, target_criterion_num):
            if(float(criterion_user_input[target_criterion[j]][1]) >= 0):
                if(float(file_textbox[i][target_criterion[j]+1]) <= float(criterion_user_input[target_criterion[j]][1])):
                    file_row_pass = False
                    break
            else:
                if(float(file_textbox[i][target_criterion[j]+1]) >= float(criterion_user_input[target_criterion[j]][1])):
                    file_row_pass = False
                    break
        if(file_row_pass == True):
            col = 1
            ws.cell(row = cur_row, column = 1, value = IO[:-4])
            for k in range(0,9):
                ws.cell(row = cur_row, column = col+1, value = file_textbox[i][col-1])
                col += 1
            cur_row += 1
    work_item = work_item + 1
# store as tmp xls file first
wb.save('tmp.xls')

# transform xls to csv
data_xls = pd.read_excel('tmp.xls',index_col=None)
data_xls.to_csv('output.csv', encoding='ANSI',sep=',',index=False,header=None)
os.remove('tmp.xls')