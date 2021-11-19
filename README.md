# excelTransform
## Environment: python 3.7
## Environment Setup
* `pip install pandas`
* `pip install openpyxl`  
* `pip install xlrd`
## Description
This code will transform several txt files into one csv file  
txt file formats is statically defined in the code  
txt formats:  
```
+------------------+-----------------------------+  
| Confusion Matrix |          Prediction         |  
|                  +---------+---------+---------+  
|    Test Data     |    GC   |  Non-GC |    KN   |  
+--------+---------+---------+---------+---------+  
|        |    GC   |     46  |     19  |     22  |  
|        +---------+---------+---------+---------+  
|  Real  |  Non-GC |     54  |     69  |    110  |  
|        +---------+---------+---------+---------+  
|        |    KN   |     34  |     35  |     76  |  
+--------+---------+---------+---------+---------+  
```
## Execute
step 1. Put all input txt files into the directory where `excelTransform.py` is resided.  
step 2. cmd `python excelTransform.py`