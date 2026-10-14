import sys

sys.path.append(r"..\PSO\ServisPSO\venv\Lib\site-packages\openpyxl") #C:\Users\Admin\PSO\venv
#sys.path.append("../PSO/ServisPSO/venv/Lib/site-packages/openpyxl")
#import openpyxl

from openpyxl import load_workbook
#import pandas as pd
from openpyxl import Workbook

from openpyxl.styles import Border, Side, Alignment, Font # для таблицы

# path to file
file_path = r'C:\Users\Fomenko.SM\PSO\ServisPSO\Noyabrsk\Справка_ППД_Noyabrsk.xlsx'

wb=load_workbook(file_path)
print("Листы:", wb.sheetnames)

ws=wb.active
#print(sheet)

# print('Содержимое всего листа:') # min_row=2 — если нужно пропустить строку с заголовками 
# for row in ws.iter_rows(min_row=1, values_only=True): # values_only=True возвращает только чистые значения (строки, числа), а не объекты ячеек 
#     # row — это кортеж значений ячеек, например (None, 'Товар 1', 500, None)
#     # all() вернет True, только если ВСЕ значения в строке равны None
#     # not all() пропустит строку, только если она полностью пустая
#     if not all(cell is None for cell in row):
#         print(row) 

print("----------------------------------------------------------------")
# Значение, которое нужно найти
search_value = 'Работы по  сервисному обслуживанию'  # Учитывайте пробелы и регистр

# Перебираем строки и столбцы
num=int()
for row in range(1, ws.max_row + 1):
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value == search_value:
            # Ячейка найдена — выводим координаты
            print(f"Значение '{search_value}' найдено в ячейке: {cell.coordinate}")
            a=cell.coordinate
            print(type(a))
            digits = ''.join(char for char in a if char.isdigit())
            num=int(digits)+5
            print(digits)
            # Дальнейшая работа с ячейкой...
            break  # Прекращаем поиск, если нужно найти первое совпадение
print('num: ',num)

tr='Работы по текущему ремонту'
num_tr=int()
for row in range(1, ws.max_row + 1):
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value == tr:
            # Ячейка найдена — выводим координаты
            print(f"Значение '{tr}' найдено в ячейке: {cell.coordinate}")
            a=cell.coordinate
            print(type(a))
            digits = ''.join(char for char in a if char.isdigit())
            num_tr=int(digits)
            print(digits)
            # Дальнейшая работа с ячейкой...
            break  # Прекращаем поиск, если нужно найти первое совпадение
print('num_tr: ',num_tr)

# Теперь работаем с данными от 9-й строки до 33-й
for row in ws.iter_rows(min_row=num, max_row=num_tr, values_only=True):
    # row — это кортеж значений текущей строки
    #print(row)  # или обрабатывайте данные как нужно
    value = row[4]
    if value is not None:          # пропускаем пустые ячейки
            print(value)
    #print(value)
wb.close()