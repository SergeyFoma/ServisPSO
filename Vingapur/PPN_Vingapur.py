import sys

sys.path.append(r"..\PSO\ServisPSO\venv\Lib\site-packages\openpyxl") #C:\Users\Admin\PSO\venv
#sys.path.append("../PSO/ServisPSO/venv/Lib/site-packages/openpyxl")
#import openpyxl

from openpyxl import load_workbook
#import pandas as pd
from openpyxl import Workbook

from openpyxl.styles import Border, Side, Alignment, Font # для таблицы

# path to file
file_path = r'C:\Users\Fomenko.SM\PSO\ServisPSO\Vingapur\План 10 Н.xlsx'

# start_PPD=int()
# start_PPN=int() # начало ППН в файле План 10
# end_PPN=int() # конец ППН в файле План 10
# sheet_name = None  # None = активный лист; можно указать имя, например "Лист1"
# cols=[1,3,6,7]

# Читаем все листы документа
# 1 - Узнать названия всех листов в документе
wb = load_workbook(file_path, data_only=True)  # data_only=True — брать значения формул, а не сами формулы
print("Листы:", wb.sheetnames)

# 2 - Выбрать нужный лист
# Пятый лист по порядку (индексация с 0, значит индекс 4) НВынгапур ТОРО
ws = wb.worksheets[4]
print("Имя листа:", ws.title)

# 3 - Прочитать весь лист
PPD = []
PPN = []

start_PPD = None
end_PPD = None
start_PPN = None
end_PPN = None

# Предположим, что метки (ППД, ППН) в столбце A (col=1), а оборудование в столбце B (col=2)
col_label = 2
col_eq = 3

rows = list(ws.iter_rows(min_col=col_label, max_col=max(col_label, col_eq), values_only=True))

for i, row in enumerate(rows):
    label = row[0] if len(row) > 0 else None
    eq = row[1] if len(row) > 1 else None

    # Нормализуем: убираем пробелы, приводим к верхнему регистру для сравнения
    if isinstance(label, str):
        label_norm = label.strip().upper()
    else:
        label_norm = None

    if label_norm == 'ППД':
        if start_PPD is None:
            start_PPD = i + 1  # номер строки (1-based)
        # Если уже был ППД и теперь снова ППД — можно либо перезаписать, либо обработать ошибку.
        # Здесь просто перезаписываем начало.

    elif label_norm == 'ППН':
        # Конец ППД — строка перед ППН
        if start_PPD is not None and end_PPD is None:
            end_PPD = i  # строка с ППН — это уже не часть ППД
        # Начало ППН — строка с меткой
        if start_PPN is None:
            start_PPN = i + 1

    # Собираем оборудование, если мы внутри нужного диапазона
    # Для ППД: от start_PPD до end_PPD (не включая строку с меткой ППН)
    if start_PPD is not None and eq is not None:
        row_num = i + 1
        if end_PPD is None:
            # Пока не встретили ППН, считаем всё после ППД частью ППД
            if row_num > start_PPD:
                PPD.append(eq)
        else:
            # После появления ППН: ППД только до end_PPD
            if start_PPD < row_num <= end_PPD:
                PPD.append(eq)

    # Для ППН: от start_PPN до конца (или до следующей метки, если будет)
    if start_PPN is not None and eq is not None:
        row_num = i + 1
        # Пока нет конца ППН, собираем всё после метки
        if row_num > start_PPN:
            PPN.append(eq)

# Если ППН не было, но был ППД, то конец ППД — последняя строка
if start_PPD is not None and end_PPD is None:
    end_PPD = len(rows)

# Если был ППН, то конец ППН — последняя строка (можно уточнить логику, если есть ещё метки)
if start_PPN is not None and end_PPN is None:
    end_PPN = len(rows)

print('PPD:', PPD)
print('PPN:', PPN)
print('start_PPD:', start_PPD, 'end_PPD:', end_PPD)
print('start_PPN:', start_PPN, 'end_PPN:', end_PPN)

# ------------------------------------------------------------------------------------------
# Сервис

rows_data = [] 
rows_data_tr=[]
rows_data_dem=[]
rows_data_mon=[]
rows_data_pnr=[]

# --- ЭТАП 1: сбор данных ---
for row in ws.iter_rows(min_row=start_PPN + 1, max_row=end_PPN, values_only=True):
    if row[7] == 1:
        rows_data.append([1, 'ЦППН-3', row[2], 'насос', row[4], f'НА-{row[3]} СО', row[5], row[1]])
        # print(row)
    if row[8] == 1:
        rows_data_tr.append([1, 'ЦППН-3', row[2], 'насос', row[4], f'НА-{row[3]} ТР', row[5], row[1]])
    if row[9] == 1:
        rows_data_dem.append(list([1, 'ЦППН-3',row[2], f'НА-{row[3]}', row[5],'', row[4], row[1]]))
    if row[10] == 1 and row[11] == 1:
        rows_data_dem.append(list([1, 'ЦППН-3',row[2], f'НА-{row[3]}', '',row[5], row[4], row[1]]))
    if row[11] == 1 and row[10] != 1:
        rows_data_dem.append(list([1, 'ЦППН-3',row[2], f'НА-{row[3]}', '',f"(ПНР){row[5]}", row[4], row[1]]))

print("len(rows_data): " ,len(rows_data))
print("ROWS_DATA: ", rows_data)

# --- ЭТАП 2: запись в файл ---

# 1. Создаём новую книгу
wb = Workbook()

# 2. Получаем активный лист (по умолчанию он уже есть)
ws = wb.active
ws.title = "Справка ППН Вынгапур"  # можно переименовать лист
# Сохраняем файл с нужным названием
wb.save(r"C:\Users\Fomenko.SM\PSO\ServisPSO\Vingapur\Справка_ППН_Vingapur.xlsx")


# -----------------------------------------------------------------------

# # 3. Пишем данные
spravka_PPN = r'C:\Users\Fomenko.SM\PSO\ServisPSO\Vingapur\Справка_ППН_Vingapur.xlsx'
wb = load_workbook(filename=spravka_PPN)
ws = wb.active

# Шапка файла
col=6 # колонка в которой будет объединение ячеек(№ п/п)
# ws.cell(row=2, column=col, value='ООО "ЯмалСпецЦентр"')
# ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col)
top_left = ws.cell(row=2, column=6)
top_left.value = 'ООО "ЯмалСпецЦентр"'
top_left.font = Font(bold=True, size=12, color='0000FF')
top_left.alignment = Alignment(
    horizontal='center',      # по горизонтали
    vertical='center',        # по вертикали
    wrap_text=True            # перенос текста
)
ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col)

col=6 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=3, column=col, value='(Наименование организации)')
ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col)

# col=6 
# ws.cell(row=4, column=col, value='СПРАВКА №04')
# ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col)

top_left = ws.cell(row=4, column=6)
top_left.value = 'СПРАВКА №04'
top_left.font = Font(bold=True, size=12)
top_left.alignment = Alignment(
    horizontal='center',      # по горизонтали
    vertical='center',        # по вертикали
    wrap_text=True            # перенос текста
)
ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col)

ws.cell(row=5, column=5, value=' о выполнении сервисных работ за')
ws.merge_cells(start_row=5, start_column=5, end_row=5, end_column=6)

# col=7 
# ws.cell(row=5, column=col, value='сентябрь')
# ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col)

top_left = ws.cell(row=5, column=7)
top_left.value = 'октябрь'
top_left.font = Font(bold=True, size=12)
top_left.alignment = Alignment(
    horizontal='center',      # по горизонтали
    vertical='center',        # по вертикали
    wrap_text=True            # перенос текста
)
ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col)

col=8 
ws.cell(row=5, column=col, value='2026 года')
ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col)

ws.cell(row=6, column=4, value=' на объектах ЦППН-1 Акционерного Общества "Газпромнефть-Ноябрьскнефтегаз"(Ноябрьский регион)')
ws.merge_cells(start_row=6, start_column=4, end_row=6, end_column=9)

# -------------------Header table for Servis------------------

# Servis таблица header строки - 9x12
start_header_so=9# строка с которой начнётя объединение ячеек
end_header_so=start_header_so+3 # количество объединённых ячеек
print("Start_header_so: ", start_header_so)
print("End_header_so: ", end_header_so)

col=1 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_so, column=col, value='№ п/п')
ws.merge_cells(start_row=start_header_so, start_column=col, end_row=end_header_so, end_column=col)

col=2 
ws.cell(row=start_header_so, column=col, value='ЦЕХ')
ws.merge_cells(start_row=start_header_so, start_column=col, end_row=end_header_so, end_column=col)

col=8 # колонка в которой будет объединение ячеек(Дата выполнения работ)
ws.cell(row=start_header_so, column=col, value='Дата выполнения работ')
ws.merge_cells(start_row=start_header_so, start_column=col, end_row=end_header_so, end_column=col)

#print("COL: ", col)
col=9 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_so, column=col, value='Подпись, ФИО механика (мастера)')
ws.merge_cells(start_row=start_header_so, start_column=col, end_row=end_header_so, end_column=col)
col=10
ws.cell(row=start_header_so, column=col, value='№ ЗАКАЗА САП ТОРО')
ws.merge_cells(start_row=start_header_so, start_column=col, end_row=end_header_so, end_column=col)

# # Объединяем ячейки для заголовка
# # ws.merge_cells('C20:G21')
end_row=start_header_so+1
start_col=3
end_col=7
# ws.cell(row=start_row, column=start_col, value='Работы по  сервисному обслуживанию ')
# ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)

top_left = ws.cell(row=start_header_so, column=start_col)
top_left.value = 'Работы по  сервисному обслуживанию'
top_left.font = Font(bold=True, size=12, color='0000FF')
top_left.alignment = Alignment(
    horizontal='center',      # по горизонтали
    vertical='center',        # по вертикали
    wrap_text=True            # перенос текста
)
ws.merge_cells(start_row=start_header_so, start_column=start_col, end_row=end_row, end_column=end_col)

start = end_row+1
end=start+1
col=3
ws.cell(row=start, column=col, value='Объект')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

# start = end_row+1
# end=start+1
col=4
ws.cell(row=start, column=col, value='Оборудование')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

# start = end_row+1
# end=start+1
col=5
ws.cell(row=start, column=col, value='Марка оборудования')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

# start = end_row+1
# end=start+1
col=6
ws.cell(row=start, column=col, value='Рег. №')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

# start = end_row+1
# end=start+1
col=7
ws.cell(row=start, column=col, value='Зав №')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

# Запись Сервис (начиная со строки 14)
cols = [1, 2, 3, 4, 5, 6, 7, 10]
for r_idx, row_data in enumerate(rows_data, start=end_header_so+2): # start_PPD+len(rows_data)):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value


# Заполняем столбец №п/п
# Настраиваем ключевой столбец, по которому считаем строку заполненной
# Например, столбец B (индекс 1)
key_col_idx = 1

start_row = end_header_so+2 #14  # пропускаем заголовок (строка 1)
filled_rows = []

# Проходим по строкам и ищем непустые ячейки в ключевом столбце
for row_num in range(start_row, ws.max_row + 1):
    cell = ws.cell(row=row_num, column=key_col_idx + 1)  # column=2 → B
    if cell.value is not None and str(cell.value).strip() != "":
        filled_rows.append(row_num)

# Теперь заполняем столбец A номерами по порядку
for idx, row_num in enumerate(filled_rows, start=1):
    ws.cell(row=row_num, column=1).value = idx

# print(f"Заполненных строк: {len(filled_rows)}")

# сервис записан в таблицу

# # ------------------------------------------------------------------------
# ------------------------ТР---------------------------------
# header tr
start_header_tr=len(rows_data)+16
end_header_tr=start_header_tr+3
print("Start_header_tr: ", start_header_tr)
print("End_header_tr: ", end_header_tr)
# Запись ППН ТР (начиная со строки 26)

col=1 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_tr, column=col, value='№ п/п')
ws.merge_cells(start_row=start_header_tr, start_column=col, end_row=end_header_tr, end_column=col)

col=2 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_tr, column=col, value='ЦЕХ')
ws.merge_cells(start_row=start_header_tr, start_column=col, end_row=end_header_tr, end_column=col)


start_col=3
end_col=7
top_left = ws.cell(row=start_header_tr, column=start_col)
top_left.value = 'Работы по текущему ремонту'
top_left.font = Font(bold=True, size=12, color='0000FF')
top_left.alignment = Alignment(
    horizontal='center',      # по горизонтали
    vertical='center',        # по вертикали
    wrap_text=True            # перенос текста
)
ws.merge_cells(start_row=start_header_tr, start_column=start_col, end_row=end_header_tr-2, end_column=end_col)

col=3
ws.cell(row=start_header_tr+2, column=col, value='Объект')
ws.merge_cells(start_row=start_header_tr+2, start_column=col, end_row=end_header_tr, end_column=col)

col=4
ws.cell(row=start_header_tr+2, column=col, value='Оборудование')
ws.merge_cells(start_row=start_header_tr+2, start_column=col, end_row=end_header_tr, end_column=col)

col=5
ws.cell(row=start_header_tr+2, column=col, value='Марка оборудования')
ws.merge_cells(start_row=start_header_tr+2, start_column=col, end_row=end_header_tr, end_column=col)

col=6
ws.cell(row=start_header_tr+2, column=col, value='Рег. №')
ws.merge_cells(start_row=start_header_tr+2, start_column=col, end_row=end_header_tr, end_column=col)

col=7
ws.cell(row=start_header_tr+2, column=col, value='Зав №')
ws.merge_cells(start_row=start_header_tr+2, start_column=col, end_row=end_header_tr, end_column=col)

col=8 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_tr, column=col, value='Дата выполнения работ')
ws.merge_cells(start_row=start_header_tr, start_column=col, end_row=end_header_tr, end_column=col)

col=9 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_tr, column=col, value='Подпись, ФИО механика (мастера)')
ws.merge_cells(start_row=start_header_tr, start_column=col, end_row=end_header_tr, end_column=col)

col=10
ws.cell(row=start_header_tr, column=col, value='№ ЗАКАЗА САП ТОРО')
ws.merge_cells(start_row=start_header_tr, start_column=col, end_row=end_header_tr, end_column=col)

# Запись ТР в таблицу

cols = [1, 2, 3, 4, 5, 6, 7, 10]
for r_idx, row_data in enumerate(rows_data_tr, start=end_header_tr+2): # start_PPD+len(rows_data)):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

# Заполняем столбец №п/п
# Настраиваем ключевой столбец, по которому считаем строку заполненной
# Например, столбец B (индекс 1)
key_col_idx = 1

# start_row = 90  # пропускаем заголовок (строка 1)
filled_rows = []

# Проходим по строкам и ищем непустые ячейки в ключевом столбце
for row_num in range(end_header_tr+2, ws.max_row + 1):
    cell = ws.cell(row=row_num, column=key_col_idx + 1)  # column=2 → B
    if cell.value is not None and str(cell.value).strip() != "":
        filled_rows.append(row_num)

# Теперь заполняем столбец A номерами по порядку
for idx, row_num in enumerate(filled_rows, start=1):
    ws.cell(row=row_num, column=1).value = idx


########-------------------------------------DEM MON PNR------------------------------

start_header_dem = end_header_tr+len(rows_data_tr)+3
end_header_dem = start_header_dem+1

col=1 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_dem, column=col, value='№ п/п')
ws.merge_cells(start_row=start_header_dem, start_column=col, end_row=end_header_dem+2, end_column=col) 
col=2 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_header_dem, column=col, value='ЦЕХ')
ws.merge_cells(start_row=start_header_dem, start_column=col, end_row=end_header_dem+2, end_column=col)

# Сначала работаем с верхней левой ячейкой объединённого диапазона
top_left = ws.cell(row=start_header_dem, column=start_col)
top_left.value = 'Работы по замене агрегатов'
top_left.font = Font(bold=True, size=12, color='0000FF')
ws.merge_cells(start_row=start_header_dem, start_column=start_col, end_row=end_header_dem, end_column=end_col)

col=3
ws.cell(row=start_header_dem+2, column=col, value='Объект')
ws.merge_cells(start_row=start_header_dem+2, start_column=col, end_row=end_header_dem+2, end_column=col)

col=4
ws.cell(row=start_header_dem+2, column=col, value='Рег. №')
ws.merge_cells(start_row=start_header_dem+2, start_column=col, end_row=end_header_dem+2, end_column=col)

col=5
ws.cell(row=start_header_dem+2, column=col, value='Демонтаж')
ws.merge_cells(start_row=start_header_dem+2, start_column=col, end_row=end_header_dem+2, end_column=col)

col=6
ws.cell(row=start_header_dem+2, column=col, value='Монтаж/ПНР')
ws.merge_cells(start_row=start_header_dem+2, start_column=col, end_row=end_header_dem+2, end_column=col)

col=7
ws.cell(row=start_header_dem+2, column=col, value='Марка оборудования')
ws.merge_cells(start_row=start_header_dem+2, start_column=col, end_row=end_header_dem+2, end_column=col)

col=8 
ws.cell(row=start_header_dem, column=col, value='Дата выполнения работ')
ws.merge_cells(start_row=start_header_dem, start_column=col, end_row=end_header_dem+2, end_column=col)

col=9 
ws.cell(row=start_header_dem, column=col, value='Подпись, ФИО механика (мастера)')
ws.merge_cells(start_row=start_header_dem, start_column=col, end_row=end_header_dem+2, end_column=col)

col=10
ws.cell(row=start_header_dem, column=col, value='№ ЗАКАЗА САП ТОРО')
ws.merge_cells(start_row=start_header_dem, start_column=col, end_row=end_header_dem+2, end_column=col)

cols = [1, 2, 3, 4, 5, 6, 7, 10]
# Запись ППД демонтаж 

for r_idx, row_data in enumerate(rows_data_dem, start=end_header_dem+4):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

# Запись ППД монтаж

for r_idx, row_data in enumerate(rows_data_mon, start=end_header_dem+4):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

# Заполняем столбец №п/п
# Настраиваем ключевой столбец, по которому считаем строку заполненной
# Например, столбец B (индекс 1)
key_col_idx = 1
filled_rows = []

# Проходим по строкам и ищем непустые ячейки в ключевом столбце
for row_num in range(end_header_dem+4, ws.max_row + 1):
    cell = ws.cell(row=row_num, column=key_col_idx + 1)  # column=2 → B
    if cell.value is not None and str(cell.value).strip() != "":
        filled_rows.append(row_num)

# Теперь заполняем столбец A номерами по порядку
for idx, row_num in enumerate(filled_rows, start=1):
    ws.cell(row=row_num, column=1).value = idx


wb.save(spravka_PPN)

