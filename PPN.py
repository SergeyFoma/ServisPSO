import sys
from openpyxl.styles import Alignment, Font

sys.path.append(r"..\PSO\ServisPSO\venv\Lib\site-packages\openpyxl") #C:\Users\Admin\PSO\venv
#sys.path.append("../PSO/ServisPSO/venv/Lib/site-packages/openpyxl")
import openpyxl

from openpyxl import load_workbook
#import pandas as pd
from openpyxl import Workbook

from openpyxl.styles import Border, Side # для таблицы


# path to file
file_path = r'C:\Users\Admin\PSO\ServisPSO\План 10 Н.xlsx'
# start_PPD='ППД'
#start_PPD=21 
# end_PPD='ППН'
#end_PPD=int()
start_PPN=int() # начало ППН в файле План 10
end_PPN=int() # конец ППН в файле План 10
sheet_name = None  # None = активный лист; можно указать имя, например "Лист1"
cols=[1,3,6,7]

# Читаем все листы документа
# 1 - Узнать названия всех листов в документе
wb = load_workbook(file_path, data_only=True)  # data_only=True — брать значения формул, а не сами формулы
print("Листы:", wb.sheetnames)

# 2 - Выбрать нужный лист
# Четвёртый лист по порядку (индексация с 0, значит индекс 3) Ноябрьск ТОРО
ws = wb.worksheets[3]
print("Имя листа:", ws.title)

# 3 - Прочитать весь лист
# Прочитать всё как список списков
data = []
for row in ws.iter_rows(values_only=True):
    data.append(list(row))
# Найти строку с ППД и индекс строки
for i in data:
    for ii in i:  
        if ii == 'ППН':
            start_PPN= data.index(i)
            end_PPN=data.index(data[len(data)-1])

# row_end=end_PPN-start_PPN
print("Start_PPN: ", start_PPN)
print("End_PPN: ", end_PPN)
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
        rows_data.append([1, 'ЦППД-1', row[2], 'насос', row[4], f'НА-{row[3]} СО', row[5], row[1]])
        # print(row)
    if row[8] == 1:
        rows_data_tr.append([1, 'ЦППД-1', row[2], 'насос', row[4], f'НА-{row[3]} ТР', row[5], row[1]])
    if row[9] == 1:
        rows_data_dem.append(list([1, 'ЦППД-1',row[2], f'НА-{row[3]}', row[5],'', row[4], row[1]]))

print("len(rows_data): " ,len(rows_data))
# --- ЭТАП 2: запись в файл ---

# 1. Создаём новую книгу
wb = Workbook()

# 2. Получаем активный лист (по умолчанию он уже есть)
ws = wb.active
ws.title = "Справка"  # можно переименовать лист
# Сохраняем файл с нужным названием
wb.save(r"C:\Users\Admin\PSO\ServisPSO\ППН_Н.xlsx")

# -----------------------------------------------------------------------

# # 3. Пишем данные
spravka_PPN = r'C:\Users\Admin\PSO\ServisPSO\ППН_Н.xlsx'
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

#-----------------------------------------------------------
    # Servis
start_SO=9# строка с которой начнётя объединение ячеек
#start_H=22
#print('start_A: ',type(start_A), start_A)
end_SO=start_SO+3 # количество объединённых ячеек
#print('end_A: ', type(end_A), end_A)
col=1 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_SO, column=col, value='№ п/п')
ws.merge_cells(start_row=start_SO, start_column=col, end_row=end_SO, end_column=col)

col=2 
ws.cell(row=start_SO, column=col, value='ЦЕХ')
ws.merge_cells(start_row=start_SO, start_column=col, end_row=end_SO, end_column=col)

col=8 # колонка в которой будет объединение ячеек(Дата выполнения работ)
ws.cell(row=start_SO, column=col, value='Дата выполнения работ')
ws.merge_cells(start_row=start_SO, start_column=col, end_row=end_SO, end_column=col)

#print("COL: ", col)
col=9 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_SO, column=col, value='Подпись, ФИО механика (мастера)')
ws.merge_cells(start_row=start_SO, start_column=col, end_row=end_SO, end_column=col)
col=10
ws.cell(row=start_SO, column=col, value='№ ЗАКАЗА САП ТОРО')
ws.merge_cells(start_row=start_SO, start_column=col, end_row=end_SO, end_column=col)

wb.save(spravka_PPN)
print("Start_SO: ", start_SO)
print("END_SO: ", end_SO)
# # Объединяем ячейки для заголовка
# # ws.merge_cells('C20:G21')
start_row=start_SO
end_row=start_row+1
start_col=3
end_col=7
# ws.cell(row=start_row, column=start_col, value='Работы по  сервисному обслуживанию ')
# ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)

top_left = ws.cell(row=start_row, column=start_col)
top_left.value = 'Работы по  сервисному обслуживанию'
top_left.font = Font(bold=True, size=12, color='0000FF')
top_left.alignment = Alignment(
    horizontal='center',      # по горизонтали
    vertical='center',        # по вертикали
    wrap_text=True            # перенос текста
)
ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)

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
wb.save(spravka_PPN)


cols = [1, 2, 3, 4, 5, 6, 7, 10]


# start_PPN_tr=start_PPN+len(rows_data)+len(rows_data_tr)+8+3
# Запись Сервис (начиная со строки 14)
for r_idx, row_data in enumerate(rows_data, start=end_SO+2): # start_PPD+len(rows_data)):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

# Заполняем столбец №п/п
# Настраиваем ключевой столбец, по которому считаем строку заполненной
# Например, столбец B (индекс 1)
key_col_idx = 1

start_row = 14  # пропускаем заголовок (строка 1)
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
# # ------------------------------------------------------------------------

start_PPN_tr=end_SO+len(rows_data)
# Запись ППД ТР (начиная со строки 26)
for r_idx, row_data in enumerate(rows_data_tr, start=start_PPN_tr-2):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value
print("Start_PPN_tr: ", start_PPN_tr)
# Заполняем столбец №п/п
# Настраиваем ключевой столбец, по которому считаем строку заполненной
# Например, столбец B (индекс 1)
key_col_idx = 1

start_row = start_PPN_tr-2  # пропускаем заголовок (строка 1)
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
# # ------------------------------------------------------------------------

cols = [1, 2, 3, 4, 5, 6, 7, 10]
# Запись ППД демонтаж (начиная со строки 46)
start_PPN_dem=start_PPN_tr+10
print("Start_PPN_dem: ", start_PPN_dem)
for r_idx, row_data in enumerate(rows_data_dem, start=start_PPN_dem):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

# Заполняем столбец №п/п
# Настраиваем ключевой столбец, по которому считаем строку заполненной
# Например, столбец B (индекс 1)
key_col_idx = 1

start_row = 90  # пропускаем заголовок (строка 1)
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
# # ------------------------------------------------------------------------

# print("Start_PPD: ", start_PPD+len(rows_data))
# print("Start_PPD_tr: ",start_PPD_tr)

# # --- СОХРАНЯЕМ ---
wb.save(spravka_PPN)


# # ------------------------------------Объединение ячеек. Формирование шапки таблицы TR-----------------------------------
#spravka_PPD = r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
wb = load_workbook(spravka_PPN)
ws = wb.active
start_H=start_PPN+6+len(rows_data)+3# строка с которой начнётя объединение ячеек
#start_H=22
#print('start_A: ',type(start_A), start_A)
end_H=start_H+3 # количество объединённых ячеек
#print('end_A: ', type(end_A), end_A)
col=1 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='№ п/п')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)

col=2 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='ЦЕХ')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)

col=8 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='Дата выполнения работ')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)

#print("COL: ", col)
col=9 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='Подпись, ФИО механика (мастера)')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)
col=10
ws.cell(row=start_H, column=col, value='№ ЗАКАЗА САП ТОРО')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)

wb.save(spravka_PPN)
# Объединяем ячейки для заголовка
# ws.merge_cells('C20:G21')
start_row=start_H
end_row=start_row+1
start_col=3
end_col=7
# ws.cell(row=start_row, column=start_col, value='Работы по текущему ремонту!!!')
# ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)

top_left = ws.cell(row=start_row, column=start_col)
top_left.value = 'Работы по текущему ремонту'
top_left.font = Font(bold=True, size=12, color='0000FF')
top_left.alignment = Alignment(
    horizontal='center',      # по горизонтали
    vertical='center',        # по вертикали
    wrap_text=True            # перенос текста
)
ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)

start = end_row+1
end=start+1
col=3
ws.cell(row=start, column=col, value='Объект')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
col=4
ws.cell(row=start, column=col, value='Оборудование')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
col=5
ws.cell(row=start, column=col, value='Марка оборудования')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
col=6
ws.cell(row=start, column=col, value='Рег. №')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
col=7
ws.cell(row=start, column=col, value='Зав №')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)
wb.save(spravka_PPN)
# # # -------------------------------------------------------------------------------
# таблица для демонтаж, монтаж, пнр
start_H=start_PPN+8+len(rows_data)+len(rows_data_tr)+8# строка с которой начнётя объединение ячеек
print('start_H: ',type(start_H), start_H)

end_H=start_H+3 # количество объединённых ячеек
#print('end_A: ', type(end_A), end_A)
col=1 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='№ п/п')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col) 
col=2 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='ЦЕХ')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)
col=8 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='Дата выполнения работ')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)
#print("COL: ", col)
col=9 # колонка в которой будет объединение ячеек(№ п/п)
ws.cell(row=start_H, column=col, value='Подпись, ФИО механика (мастера)')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)
col=10
ws.cell(row=start_H, column=col, value='№ ЗАКАЗА САП ТОРО')
ws.merge_cells(start_row=start_H, start_column=col, end_row=end_H, end_column=col)


# Объединяем ячейки для заголовка
# ws.merge_cells('C20:G21')
start_row=14+len(rows_data)+len(rows_data_tr)+10+17
end_row=start_row+1
start_col=3
end_col=7
print("Start_row: ", start_row)


# Сначала работаем с верхней левой ячейкой объединённого диапазона
top_left = ws.cell(row=start_row, column=start_col)
top_left.value = 'Работы по замене агрегатов'
top_left.font = Font(bold=True, size=12, color='0000FF')
ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)

start = end_row+1
end=start+1
col=3
ws.cell(row=start, column=col, value='Объект')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
col=4
ws.cell(row=start, column=col, value='Рег. №')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
col=5
ws.cell(row=start, column=col, value='Демонтаж')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
col=6
ws.cell(row=start, column=col, value='Монтаж/ПНР')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

start = end_row+1
end=start+1
#print("END: ", end)
col=7
ws.cell(row=start, column=col, value='Марка оборудования')
ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)



# # Записываем текст в левую верхнюю ячейку
# #ws['C20'] = 'Работы по текущему ремонту!!!'

# # Можно дополнительно оформить (например, центрировать текст)
# #ws['C20'].alignment = Alignment(horizontal='center', vertical='center')


wb.save(spravka_PPN)
# # # Явное закрытие
# # wb.close()
# print("Данные записаны.")
# #print('rows_data_tr: ', rows_data_tr)
# #print('len(rows_data): ',len(rows_data))

