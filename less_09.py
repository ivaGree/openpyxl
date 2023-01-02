# less9.py - Вариант 3 (for row in ws.values): Читаем лист из Excel файла

# Подключили openpyxl
import openpyxl
from openpyxl import *

# Подключаем эксельку с нашей базой
wb = load_workbook("database.xlsx")

# Постучаться к листу (staff) можно одним из 3х способов / <Worksheet "staff">
# ws = wb.active        # 1) по активному листу
# ws = wb.worksheets[0] # 2) по порядковому номеру в книге, если в файле поменять местами листы - то все поломается
ws = wb['staff']        # 3) по конкретному названию листа в книге

for row in ws.values:
    for cell in row:
        print(cell)