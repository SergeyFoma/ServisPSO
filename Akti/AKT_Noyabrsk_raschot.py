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
num_so=int()
for row in range(1, ws.max_row + 1):
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value == search_value:
            # Ячейка найдена — выводим координаты
            print(f"Значение '{search_value}' найдено в ячейке: {cell.coordinate}")
            a=cell.coordinate
            print(type(a))
            digits = ''.join(char for char in a if char.isdigit())
            num_so=int(digits)+5
            #print(digits)
            # Дальнейшая работа с ячейкой...
            break  # Прекращаем поиск, если нужно найти первое совпадение
print('Начинается запись со строки num_so: ',num_so)

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
            #print(digits)
            # Дальнейшая работа с ячейкой...
            break  # Прекращаем поиск, если нужно найти первое совпадение
print('num_tr: ',num_tr)
end_num_tr=num_tr+4
print('end_num_tr: ', end_num_tr)

# Чтение раздела ТР
data_tr=[]
for row in ws.iter_rows(min_row=end_num_tr+1, max_row=end_num_tr+6, values_only=True):
    # row — это кортеж значений текущей строки
    #print(row)  # или обрабатывайте данные как нужно
    value = [row[4]]
    # if value is not None:          # пропускаем пустые ячейки
    #         print('value: ',value)
    if value is None or value == '':
        break       # останавливаемся на первой пустой ячейке
    #print(value)
    data_tr.append(value)
print(data_tr)
wb.close()


#----------------------------------------------------------------

#Записываем в акт ТР
# Путь к файлу и имя листа
file_path = r"C:\Users\Fomenko.SM\PSO\ServisPSO\Noyabrsk\Акт факт ННГ-22-10000-02048-Р октябрь.xlsx"
sheet_name = "Акт июнь"  # или другое имя листа

# Загружаем книгу в режиме редактирования (не только для чтения)
wb = load_workbook(file_path)
ws = wb[sheet_name]


import re

# cns='НЦС', 'ЦНС', 'нцс', 'цнс'

# pattern_240 = re.compile(r'\bЦНС\s+240\b')  # границы слова вокруг ЦНС и 240
cns_base = "ЦНС"  # одна строка, регистр не важен из-за флага IGNORECASE
row=145
pattern_240 = re.compile(rf'\b{re.escape(cns_base)}\s+240\b', re.IGNORECASE)

matched = [
    item[0]
    for item in data_tr
    if item and item[0] and pattern_240.search(item[0])
]
print('matched: ', matched, len(matched))

# Записываем значение в ячейку D157 (строка 157, столбец 4)
value_to_write = len(matched)
ws.cell(row=row, column=4, value=value_to_write)
# Сохраняем изменения
wb.save(file_path)

# -----------------------------------------------------------------------------
cns_base = "НЦС"  # одна строка, регистр не важен из-за флага IGNORECASE

pattern_320 = re.compile(rf'\b{re.escape(cns_base)}\s+320\b', re.IGNORECASE)

matched = [
    item[0]
    for item in data_tr
    if item and item[0] and pattern_320.search(item[0])
]
print('matched: ', matched, len(matched))

# Записываем значение в ячейку D157 (строка 157, столбец 4)
value_to_write = len(matched)
ws.cell(row=151, column=4, value=value_to_write)

# Сохраняем изменения
wb.save(file_path)

#________________________180______________________
cns_variants = ['НЦС', 'ЦНС', 'нцс', 'цнс']
# Сначала делаем строковую часть шаблона
pattern_str = '|'.join(re.escape(v) for v in cns_variants)

# Потом компилируем в настоящий regex-объект
pattern_180 = re.compile(rf'\b({pattern_str})\s+180\b')

matched = [
    item[0]
    for item in data_tr
    if item and item[0] and pattern_180.search(item[0])
]
print('matched: ', matched, len(matched))

# Записываем значение в ячейку D157 (строка 157, столбец 4)
value_to_write = len(matched)
ws.cell(row=142, column=4, value=value_to_write)

# Сохраняем изменения
wb.save(file_path)
