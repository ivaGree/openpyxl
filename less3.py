# Less3: Запишем названия листов в переменные

import openpyxl
from openpyxl import load_workbook

# ----------------------------------------------------------------------------------------------------------------------
# Подключаем эксельку с нашей базой
wb = load_workbook("database.xlsx")

# Записываем листы в переменные
ws_staff = wb['staff']          # ws -> ws_staff
ws_worklog = wb['worklog']      # ws -> ws_worklog
print(ws_staff)                 # <Worksheet "staff">
print(type(ws_staff))           # ТИП: <class 'openpyxl.worksheet.worksheet.Worksheet'>