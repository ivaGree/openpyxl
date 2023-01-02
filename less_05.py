# less5.py - Получаем в список заголовки листа

# Подключили openpyxl
import openpyxl
from openpyxl import *

# Подключаем эксельку с нашей базой
wb = load_workbook("database.xlsx")
# Записываем название листа в переменную
ws = wb['staff']

# Определяем заголовки столбцов листа staff
rows = ws.rows
# print(rows)             # <generator object Worksheet._cells_by_row at 0x103cb5900>
# print(type(rows))       # <class 'generator'>

# Пишем заголовки столбцов в список header_list
header_list = [cell.value for cell in next(rows)]
print(header_list)       # ['ID', 'Фио', 'Логин', 'Роль']