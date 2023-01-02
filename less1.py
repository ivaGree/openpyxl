# Less1: получить активный лист книги

import openpyxl
from openpyxl import load_workbook

# ----------------------------------------------------------------------------------------------------------------------
# Подключаем эксельку с нашей базой
wb = load_workbook("database.xlsx")
#  Получить активный лист книги
ws = wb.active          # ws -> ws

print(ws)               # <Worksheet "worklog">