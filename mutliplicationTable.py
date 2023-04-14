#! python3
# mutliplicationTable.py — An exercise in working with Excel files.

import sys
import openpyxl

wb = openpyxl.Workbook()
sheet = wb.active

wb.save("multiplication_table.xlsx")
