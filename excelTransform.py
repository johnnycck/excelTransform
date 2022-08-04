import pandas as pd
import math
import os
import numpy as np
import openpyxl
from openpyxl import Workbook
import shutil

path = os.getcwd()
files = os.listdir(path)
files_xlsx = [f for f in files if f[-4:] == 'xlsx']
print('number of files: '+str(len(files_xlsx)))
# 批次處理同資料夾內的所有 csv 檔
for work_item in range (0,len(files_xlsx)):
    wb = Workbook()
    ws = wb.active
    IO = files_xlsx[work_item]
    sheet = pd.read_excel(IO)
    ws.cell(row = 1, column = 1, value = 'm/z')
    ws.cell(row = 1, column = 2, value = 'intensity')
    row_cnt = 2
    acl = sheet.values[0][1]
    acl_row = 1
    for i in range(1,len(sheet.values)):
        if (sheet.values[i][0] != sheet.values[i-1][0]):
            ws.cell(row = row_cnt, column = 1, value = sheet.values[i-1][0])
            ws.cell(row = row_cnt, column = 2, value = acl/acl_row)
            acl = sheet.values[i][1]
            acl_row = 1
            row_cnt = row_cnt + 1
        else:
            acl = acl + sheet.values[i][1]
            acl_row = acl_row + 1
    
    ws.cell(row = row_cnt, column = 1, value = sheet.values[len(sheet.values)-1][0])
    ws.cell(row = row_cnt, column = 2, value = acl/acl_row)
    # 存檔並刪除舊的 CSV
    print('\nfile number '+str(work_item+1) +' has finished')
    wb.save(IO[:-5] + '_even.xlsx')
    work_item = work_item + 1