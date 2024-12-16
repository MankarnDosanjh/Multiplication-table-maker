'''multiplication_table.py
Create a program multiplicationTable.py that takes a number N from the command line and creates an NxN multiplication table in an Excel spreadsheet.'''

# Sets working directory spreadsheet will be stored in
import os, sys
os.chdir(os.path.dirname(sys.argv[0]))

import openpyxl
from openpyxl import utils
mult = 6

wb = openpyxl.Workbook()
sheet = wb.active

# Generates table headings
for x in range(2, mult + 2):
    sheet.cell(row=x, column=1).value = x - 1
    sheet.cell(row=1, column=x).value = x - 1

max_row = sheet.max_row
max_column = utils.get_column_letter(sheet.max_column)

for cell_row in sheet['B2': f'{max_column}{max_row}']:
    for cell in cell_row:
        cur_column = utils.get_column_letter(cell.column)
        row_mult = sheet[f'A{cell.row}'].value
        col_mult = sheet[f'{cur_column}1'].value
        cell.value = row_mult * col_mult

wb.save(f'./{mult}_times_table.xlsx')