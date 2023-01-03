# less_15.py - Применим (ВПР) в рамках двух листов

import openpyxl
from openpyxl import load_workbook
import datetime as dt

# Подключаем эксельку с нашей базой
wb = load_workbook("database.xlsx")
ws_staff = wb['staff']
ws_worklog = wb['worklog']

for row in ws_staff.iter_rows(min_row=2, max_row=ws_staff.max_row, min_col=1, max_col=ws_staff.max_column):
    row = [cell.value for cell in row]
    login = row[2]

    sum = 0
    for row in ws_worklog.iter_rows(min_row=2, max_row=ws_worklog.max_row, min_col=1, max_col=ws_worklog.max_column):
        row = [cell.value for cell in row]
        login2 = row[3]
        worklog = row[1]
        if login == login2:
            sum += worklog
    print(f'{login} WORKED: {sum} hours')





