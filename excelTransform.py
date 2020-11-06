import pandas as pd
import math
import numpy as np
import openpyxl
from openpyxl import Workbook

# xls read file name
# Row 1 of input file must be empty
IO = 'test.xls'
sheet = pd.read_excel(io=IO)

# create a new excel workbook
wb = Workbook()
# select active work sheet
ws = wb.active

# Store 24 Varient Name as varienName list
varientName = ['SAA2a (RS-)', 'SAA2b (RS-)', 'SAA1g (RS-)', 'SAAU2 (RS-)', \
            'SAA1a (RS-)', 'SAAU1 (RS-)', 'SAA2a (R-)', 'SAA2b (R-)', \
            'SAA1b (RS-)', 'SAA1g (R-)', 'SAAU3 (RS-)', 'SAAU2 (R-)', \
            'SAA1a (R-)', 'SAAU1 (R-)', 'SAA1b (R-)', 'SAAU3 (R-)', \
            'SAA2a', 'SAA2b', 'SAA1g', 'SAAU2', 'SAA1a', 'SAAU1', 'SAA1b', 'SAAU3']

# assign 24 title and initialize value(0)
for i in range(1,len(sheet.values)+1):
    for j in range(0,24):
        # precess every three rows
        if ((i%3) == 1):
            # col = j+5, because first four column are titles
            ws.cell(row = i, column = j+5, value = varientName[j])
        else:
            ws.cell(row = i, column = j+5, value = 0)

# length of sheet.values represent total row numbers
for i in range(0,len(sheet.values),3):
    # '?' number
    q_num = 0 
    print('processing row: ' + str(i))
    # length of sheet.values[1] represent max column numbers
    for j in range(4,len(sheet.values[1])):
        varient_num = 0
        # process '?' title
        if(sheet.values[i][j] == '?'):
            ws.cell(row = i+1, column = q_num+29, value = '?')
            ws.cell(row = i+2, column = q_num+29, value = sheet.values[i+1][j])
            ws.cell(row = i+3, column = q_num+29, value = sheet.values[i+2][j])
            q_num = q_num+1
            #print(str(sheet.values[i][j]) + "," + str(sheet.values[i+1][j]) + "," + str(sheet.values[i+2][j]))
        # process the last title, and break the loop
        elif(sheet.values[i][j] == 'ISD:'):
            ws.cell(row = i+1, column = q_num+29+1, value = 'ISD:')
            ws.cell(row = i+2, column = q_num+29+1, value = sheet.values[i+1][j])
            ws.cell(row = i+3, column = q_num+29+1, value = sheet.values[i+2][j])
            break
        # process 24 varients
        else:
            for k in range(varient_num,24):
                if(sheet.values[i][j] == varientName[k]):
                    ws.cell(row = i+2, column = k+5, value = sheet.values[i+1][j])
                    ws.cell(row = i+3, column = k+5, value = sheet.values[i+2][j])
                    varient_num = k

# xls file name
wb.save('Trans_test.xls')