# excelTransform
## Environment: python 3.7
## Environment Setup
* `pip install pandas`
* `pip install openpyxl`  
* `pip install xlrd`
## Description
This code will transform txt files into csv file.  
Firstly, it will ask users to enter desired criterions. Users can enter **00** for don't care criterions.  
Type positive number for `>`.  
Type negative number for `<`. For instance, `TP: -5` means find the number which is less than 5.  
This version does not support `>=` or `<=`.  
For text box is `nan`, it will correspond to any criterion.  
The code will do the rest analysis, and store the result into one csv file.
## Execute
step 1. Put all input txt files into the directory where `excelTransform.py` is resided.  
step 2. cmd `python excelTransform.py`