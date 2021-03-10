# excelTransform
## Environment: python 3.7
## Environment Setup
pip install pandas  
pip install openpyxl  
pip install xlrd
## Description
This code will transform txt files into csv files. Firstly, it will check the nunber of m/z between 5000~12450. If the number exceeds 16384, it will check the number of m/z between 11300~12450. If it still exceeds 16384, this txt files will be ignored.
## Execute
step 1. Transform txt files in current directory to csv files.
step 2. python excelTransform.py