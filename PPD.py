import sys
from openpyxl.styles import Alignment, Font

sys.path.append("../PSO/ServisPSO/venv/Lib/site-packages/openpyxl") #C:\Users\Admin\PSO\venv
#sys.path.append("../PSO/ServisPSO/venv/Lib/site-packages/openpyxl")
import openpyxl

from openpyxl import load_workbook
#import pandas as pd
from openpyxl import Workbook


# path to file
file_path = r'C:\Users\Fomenko.SM\PSO\ServisPSO\План 10 Н.xlsx'
# start_PPD='ППД'
start_PPD=21
# end_PPD='ППН'
end_PPD=int()
start_PPN=int()
end_PPN=int()
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
print(data[8])
for i in data:
    for ii in i:
        if ii == 'ППД':
            start_PPD=data.index(i)+1
            #end_PPD = data.index(i)
        # найти индекс ППН     
        if ii == 'ППН':
            end_PPD = data.index(i)
            start_PPN= data.index(i)
            end_PPN=data.index(data[len(data)-1])

# print('start_PPD: ', start_PPD)
row_end=end_PPD-start_PPD
# print('row_end: ', row_end)
# print("START_PPD: ", start_PPD)
# ------------------------------------------------------------------------------------------
# Сервис

rows_data = []
rows_data_tr=[]
rows_data_dem=[]
rows_data_mon=[]
rows_data_pnr=[]

# --- ЭТАП 1: сбор данных ---
for row in ws.iter_rows(min_row=start_PPD + 1, max_row=end_PPD, values_only=True):
    if row[7] == 1:
        rows_data.append([1, 'ЦППД-1', row[2], 'насос', row[4], f'НА-{row[3]} СО', row[5], row[1]])
    if row[8] == 1:
        rows_data_tr.append([1, 'ЦППД-1', row[2], 'насос', row[4], f'НА-{row[3]} ТР', row[5], row[1]])
    if row[9] == 1:
        rows_data_dem.append(list([1, 'ЦППД-1',row[2], f'НА-{row[3]}', row[5],'', row[4], row[1]]))

# --- ЭТАП 2: запись в файл ---

# 1. Создаём новую книгу
wb = Workbook()

# 2. Получаем активный лист (по умолчанию он уже есть)
ws = wb.active
ws.title = "Справка"  # можно переименовать лист

# # 3. Пишем данные
spravka_PPD = r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
wb = load_workbook(filename=spravka_PPD)
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

col=6 
ws.cell(row=4, column=col, value='СПРАВКА №04')
ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col)

ws.cell(row=5, column=5, value=' о выполнении сервисных работ за')
ws.merge_cells(start_row=5, start_column=5, end_row=5, end_column=6)

col=7 
ws.cell(row=5, column=col, value='сентябрь')
ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col)

col=8 
ws.cell(row=5, column=col, value='2026 года')
ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col)

ws.cell(row=6, column=4, value=' на объектах ЦППН-1 Акционерного Общества "Газпромнефть-Ноябрьскнефтегаз"(Ноябрьский регион)')
ws.merge_cells(start_row=6, start_column=4, end_row=6, end_column=9)

#-----------------------------------------------------------
    # Servis
start_H=9# строка с которой начнётя объединение ячеек
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

wb.save(spravka_PPD)
# Объединяем ячейки для заголовка
# ws.merge_cells('C20:G21')
start_row=start_H
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
wb.save(spravka_PPD)


cols = [1, 2, 3, 4, 5, 6, 7, 10]


start_PPD_tr=start_PPD+len(rows_data)+len(rows_data_tr)+8+3
# Запись Сервис (начиная со строки 14)
for r_idx, row_data in enumerate(rows_data, start=14): # start_PPD+len(rows_data)):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

# Запись ППД ТР (начиная со строки 26)
for r_idx, row_data in enumerate(rows_data_tr, start=start_PPD_tr-2):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

cols = [1, 2, 3, 4, 5, 6, 7, 10]
# Запись ППД демонтаж (начиная со строки 46)
for r_idx, row_data in enumerate(rows_data_dem, start=46):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

print("Start_PPD: ", start_PPD+len(rows_data))
print("Start_PPD_tr: ",start_PPD_tr)

# --- СОХРАНЯЕМ ---
wb.save(spravka_PPD)


# ------------------------------------Объединение ячеек. Формирование шапки таблицы TR-----------------------------------
spravka_PPD = r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
wb = load_workbook(spravka_PPD)
ws = wb.active
start_H=start_PPD+6+len(rows_data)+3# строка с которой начнётя объединение ячеек
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

wb.save(spravka_PPD)
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
wb.save(spravka_PPD)
# # -------------------------------------------------------------------------------
# таблица для демонтаж, монтаж, пнр
start_H=start_PPD+8+len(rows_data)+len(rows_data_tr)+8+4+2# строка с которой начнётя объединение ячеек
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
start_row=start_PPD+8+len(rows_data)+len(rows_data_tr)+8+6
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


wb.save(spravka_PPD)
# # Явное закрытие
# wb.close()
print("Данные записаны.")
#print('rows_data_tr: ', rows_data_tr)
#print('len(rows_data): ',len(rows_data))

