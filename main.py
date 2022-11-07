import time
import subprocess
import numpy
import pandas as pd
import self as self
from openpyxl import load_workbook
import psutil as psu
import signal
import os
# forced shutdown excel to complete saving
os.system("taskkill /f /im  EXCEL.exe")

# path to Analysis_Plan_Template.xlsx file
file_path = os.path.realpath('Analysis_Plan_Template.xlsx')
print(file_path)
path_excel = file_path

# loading excel file to script and reading properly sheet
x_list = pd.read_excel(path_excel,sheet_name='Default analysis matrix')
block_type = pd.read_excel(path_excel,sheet_name='Blocks')

# Opening excel file , path to  executable file excel and xlsx file
file = path_excel
prog = r"c:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" # path to exe file of excel , flexible path
# OpenIt = subprocess.Popen([prog, file])


# names of analysis save to one list
col = list(x_list.columns.values) #columns of x list
wb = load_workbook(path_excel)
sheet = wb['Blocks']

# OpenIt.terminate()

# Algorithm
for l in range(0, block_type.shape[0]):
    list_2 = block_type[l:l + 1]["Block type"].values  ## list of block type
    for i in range(0, x_list.shape[0]):
        list_1 = x_list[i:i + 1]['Block / Analysis'].values  # list of analysis
        if list_2 == list_1: # comparison two row in Block Type and  Block / Analysis
            list_3 = x_list[i:i + 1].values
            for g in range(0, x_list.shape[1]):
                if list_3[0, g:g + 1] == 'X': # check if list have X
                    xwork = sheet.cell(row=l + 2, column=4)
                    if xwork.value:
                        xwork.value = xwork.value + "," + str(col[g])
                    else:
                        xwork.value = str(col[g])
                elif list_3[0, g:g + 1] == '?': # check if list have ?
                    qwork = sheet.cell(row=l + 2, column=5)
                    if qwork.value:
                        qwork.value = qwork.value + "," + str(col[g])
                    else:
                        qwork.value = str(col[g])
                elif g == 13: # check if all analysis is filled to one Block Type
                    wb.save(path_excel) # save workbook require permission to file , file must be closed


subprocess.Popen([prog, file])  # open file excel to show

print("EXCEL IS ALREADY FILLED")

