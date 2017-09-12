
################################################################################
# Master_file_compilation_v1
# Script for generating single Master excel file from multiple excel files
# Credit: https://stackoverflow.com/questions/15793349/how-to-concatenate-three-excels-files-xlsx-using-python
# compiled and modified script by UXO_India, 12 sep 2017
# Instructions to use:
# 1.set output folder and file name before executing script # by default it will store in C drive with Master_file_compiled.xls name
# 2. provide input workspace folder path where all excel files stored 
###################################################################################

import xlwt
import xlrd
import os #, sys,win32com.client, glob


inputWS = raw_input("Enter input workspace folder path: ") # input workspace folder
#path = os.getcwd()
path = inputWS


wkbk = xlwt.Workbook()
outsheet = wkbk.add_sheet('Locale compiled')
files = os.listdir(path)
files_xlsx = [f for f in files if f[-4:] == 'xlsx']
print files_xlsx
xlsfiles = files_xlsx
a = files_xlsx[0]

outrow_idx = 0
for f in xlsfiles:
    insheet = xlrd.open_workbook(f).sheets()[1]
    if f==a:
     for row_idx in xrange(insheet.nrows):
         for col_idx in xrange(insheet.ncols):
            outsheet.write(outrow_idx, col_idx, 
                           insheet.cell_value(row_idx, col_idx))
         outrow_idx += 1         
    else:
      for row_idx in xrange(10,insheet.nrows):
          for col_idx in xrange(insheet.ncols):
             outsheet.write(outrow_idx, col_idx, 
                           insheet.cell_value(row_idx, col_idx))
          outrow_idx += 1
        
wkbk.save(r'C:\Master_file_compiled.xls') # path and name of file
print "\n Master file successfully compiled"

###### Thank You ############


