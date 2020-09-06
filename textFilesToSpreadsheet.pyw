#! /usr/bin/env python3
# textFilesToSpreadsheet.pyw

'''
Write a program to read in the contents of several text files
and insert those contents into a spreadsheet, with one line of text per row. 
The lines of the first text file will be in the cells of column A, 
the lines of the second text file will be in the cells of column B, and so on.
'''

import openpyxl, pprint, sys, pathlib, os
from pathlib import Path

# Open new workbook
wb = openpyxl.Workbook()
# Define sheet
sheet = wb['Sheet']

# Set cwd
txtScript_dir = '/' #use the directory containing your text files
os.chdir(txtScript_dir)

# Define dir to get txt files from
txtFile_dir = './textFiles/'
# Define list to append lists of readlines() to
readlines_list = []
# Loop through each text file in folder to readlines its contents
for folderName, subfolders, filenames in os.walk(txtFile_dir):
    for filename in filenames:
        if filename[-4:] == '.txt':
            # Readlines() for each file to get list of strings
            txtFile = open(Path(txtFile_dir, filename))
            readlines = txtFile.readlines()
            # Save each list to a list variable (list of lists)
            readlines_list.append(readlines)
        else:
            continue
pprint.pprint(readlines_list)
# Loop through the list of lists to map to a new spreadsheet column
for i in range(0,len(readlines_list)):
    for j in range(0,len(readlines_list[i])):
        sheet.cell(row=j+1,column=i+1).value = readlines_list[i][j]
        
# Save workbook
wb.save('textToFile.xlsx')
