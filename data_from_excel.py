import openpyxl
from pathlib import Path

xlsx_file = Path('zxc.xlsx')

wb_obj = openpyxl.load_workbook(xlsx_file)
sheet = wb_obj.active
print(sheet)

print(sheet["A2"].value)
