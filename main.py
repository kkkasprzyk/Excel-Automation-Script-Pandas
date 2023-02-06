import time
import subprocess
import numpy
import numpy as np
import pandas as pd
import self as self
from openpyxl import load_workbook
import psutil as psu
import signal
import os

# forced shutdown excel to complete saving
# dopiska:as/stability to filter , z PI jest DC tylko , POC to net bateryjny lub konektora , dekapling nie ma tam sensu , bat, kl15,kl30 ,poc  /// notatki co i jak !
################ os.system("taskkill /f /im  EXCEL.exe")  ##########################################
os.system("taskkill /f /im  EXCEL.exe")

# path to Analysis_Plan_Template.xlsx file
file_path = os.path.realpath('Iveco_Cluster_Ticket_List_xlsx.xlsx')
file_path_2 = os.path.realpath('Analysis_Plan_Iveco_Cluster.xlsx')
test_file = os.path.realpath('testxlsx.xlsx')  ## plik testowy excela do uzupelnienia

# loading excel file to script and reading properly sheet
issues = pd.read_excel(test_file,sheet_name='Issues')
setup = pd.read_excel(file_path,sheet_name='Setup')
block_interface = pd.read_excel(file_path,sheet_name='Block | Interface')
block_type = pd.read_excel(file_path_2,sheet_name='Blocks')


# Opening excel file , path to  executable file excel and xlsx file
prog = r"c:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" # path to exe file of excel , flexible path
# OpenIt = subprocess.Popen([prog, file])


# list_2 ma miec wielkosc ilosci analiz w excelu analysis plan template czyli ilosci wierszy
list_2 =np.empty([2,block_type.shape[0]],dtype=object)
optional =np.empty([2,block_type.shape[0]],dtype=object)
priority =np.empty([2,block_type.shape[0]],dtype=object)

wb = load_workbook(file_path,data_only=True)
sheet_setup = wb['Setup']
wb_issues = load_workbook(test_file)
sheet_issues= wb_issues['Issues']

# OpenIt.terminate()


for l in range(0, block_type.shape[0]):
    list_2[0][l] = str(block_type[l:l + 1]["Block name"].values)  ## list of block type
    # print(list_2[0][l], '+')
    for s in range(1,2):
        list_2[s][l] = block_type[l:l + 1]["Default analysis list"].str.split(', ', expand=True).values.tolist()
        optional[s][l] = block_type[l:l + 1]["Optional analysis list"].str.split(', ', expand=True).values.tolist()
        priority[s][l] = block_type[l:l + 1]["Priority"].str.split(', ', expand=True).values.tolist()
        print(priority[s][l])
        # wyswietlanie ilosci analiz na jeden block name(POC,3v3)
        #   print(np.shape(list_2[1][l]),'+')
        #   print(list_2[1][0],'+')


for k in range(0, block_type.shape[0]):
        print(np.shape(list_2[1][k])[1])
        for q in range(0,np.shape(list_2[1][k])[1]):
            print(list_2[1][k][0][q])
            if list_2[1][k][0][q] == 'PI':
                print("znaleziono PI")
                x = sheet_issues.cell(row=2, column=7)
                x.value = x.value + "tutaj"
            if q == ((np.shape(list_2[1][k])[1])-1):
                print(list_2[0][k])


x = sheet_issues.cell(row=2,column=7)
x.value = x.value + "tutaj"
wb_issues.save(test_file)
#
# NOTSY# # names of analysis save to one list
# # col = list(x_list.columns.values) #columns of x list
# # wb = load_workbook(path_excel)
# # sheet = wb['Blocks']
#
#    xwork = sheet.cell(row=l + 2, column=4)
#                     if xwork.value:
#                         xwork.value = xwork.value + "," + str(col[g])
#