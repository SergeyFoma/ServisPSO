import sys

sys.path.append("../PSO/ServisPSO/venv/Lib/site-packages/openpyxl") #C:\Users\Admin\PSO\venv
#sys.path.append("../PSO/ServisPSO/venv/Lib/site-packages/openpyxl")
import openpyxl

from openpyxl import load_workbook
#import pandas as pd


# path to file
file_path = r'C:\Users\Fomenko.SM\PSO\ServisPSO\План 09 Н.xlsx'
# start_PPD='ППД'
start_PPD=int()
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
        rows_data.append([1, 'ЦППД-1', row[2], 'насос', row[4], f'НА-{row[3]}', row[5], row[1]])
    if row[8] == 1:
        rows_data_tr.append([1, 'ЦППД-1', row[2], 'насос', row[4], f'НА-{row[3]}', row[5], row[1]])
    if row[9] == 1:
        rows_data_dem.append(list([1, 'ЦППД-1',row[2], f'НА-{row[3]}', row[5],'', row[4], row[1]]))

# --- ЭТАП 2: запись в файл ---
spravka_PPD = r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
wb = load_workbook(filename=spravka_PPD)
ws = wb.active

cols = [1, 2, 3, 4, 5, 6, 7, 10]


start_PPD_tr=start_PPD+len(rows_data)+len(rows_data_tr)+8
# Запись Сервис (начиная со строки 14)
for r_idx, row_data in enumerate(rows_data, start=start_PPD+len(rows_data)):
    for col_num, value in zip(cols, row_data):
        ws.cell(row=r_idx, column=col_num).value = value

# Запись ППД ТР (начиная со строки 26)
for r_idx, row_data in enumerate(rows_data_tr, start=start_PPD_tr):
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

# #print("LEN ROWS_DATA: ", len(rows_data))
# spravka_PPD=r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
# for row in ws.iter_rows(min_row=start_PPD+1, max_row=end_PPD, values_only=True):
#     # Сервис
#     if row[7] == 1:
#         rows_data.append(list([1, 'ЦППД-1',row[2],'насос', row[4], f'НА-{row[3]}', row[5], row[1]]))#row[1]
#     # Запись данных в справку Сервис
#     #spravka_PPD=r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
#     cols = [1,2,3, 4, 5, 6, 7, 10]
#     start_row=14

#     # ВАЖНО: без read_only=True, иначе запись невозможна
#     wb = load_workbook(filename=spravka_PPD)
#     ws = wb.active

#     for r_idx, row_data in enumerate(rows_data, start=start_row):
#         for col_num, value in zip(cols, row_data):
#             ws.cell(row=r_idx, column=col_num).value = value

# # ------------------------------------------------------------------------------------  

# # Запись ППД ТР


#     if row[8] == 1:
#         rows_data_tr.append(list([1, 'ЦППД-1',row[2],'насос', row[4], f'НА-{row[3]}', row[5], row[1]]))

#     #spravka_PPD=r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
#     cols = [1,2,3, 4, 5, 6, 7, 10]
#     start_row_tr=26
#     print("START_ROW_TR: ", start_row_tr)
#     wb = load_workbook(filename=spravka_PPD)
#     ws = wb.active
#     # ВАЖНО: без read_only=True, иначе запись невозможна
#     for r_idx, row_data in enumerate(rows_data_tr, start=start_row_tr):
#         for col_num, value in zip(cols, row_data):
#             ws.cell(row=r_idx, column=col_num).value = value



# #     # -------------------------------------------------------------

# #     # Запись ППД Демонтаж

#     if row[9] == 1:
#             rows_data_dem.append(list([1, 'ЦППД-1',row[2], f'НА-{row[3]}',row[5], row[4], row[1]]))
    
#     spravka_PPD=r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
#     cols = [1,2,3, 4, 5, 7, 10]
#     start_row_dem=49
#     wb = load_workbook(filename=spravka_PPD)
#     ws = wb.active
#     # ВАЖНО: без read_only=True, иначе запись невозможна
#     for r_idx, row_data in enumerate(rows_data_dem, start=start_row_dem):
#         for col_num, value in zip(cols, row_data):
#             ws.cell(row=r_idx, column=col_num).value = value


# #     # Запись ППД Монтаж

#     if row[10] == 1:
#         rows_data_mon.append(list([1, 'ЦППД-1',row[2],f'НА-{row[3]}',row[5], row[4], row[1]]))
    
#     spravka_PPD=r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
#     cols = [1,2,3, 4, 6, 7, 10]
#     start_row_mon=49+len(rows_data_mon)
#     wb = load_workbook(filename=spravka_PPD)
#     ws = wb.active
#     # ВАЖНО: без read_only=True, иначе запись невозможна
#     for r_idx, row_data in enumerate(rows_data_mon, start=start_row_mon):
#         for col_num, value in zip(cols, row_data):
#             ws.cell(row=r_idx, column=col_num).value = value

# wb.save(spravka_PPD)
# # Явное закрытие
# wb.close()
# ------------------------------------Объединение ячеек. Формирование шапки таблицы TR-----------------------------------
spravka_PPD = r'C:\Users\Fomenko.SM\PSO\ServisPSO\ППД_Н.xlsx'
wb = load_workbook(spravka_PPD)
ws = wb.active
#start_H=start_PPD+6+len(rows_data)# строка с которой начнётя объединение ячеек
start_H=22
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
ws.cell(row=start_row, column=start_col, value='Работы по текущему ремонту!!!')
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
start_H=start_PPD+8+len(rows_data)+len(rows_data_tr)+8+4# строка с которой начнётя объединение ячеек
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


# Объединяем ячейки для заголовка
# ws.merge_cells('C20:G21')
start_row=start_PPD+8+len(rows_data)+len(rows_data_tr)+8+4
end_row=start_row+1
start_col=3
end_col=7
ws.cell(row=start_row, column=start_col, value='Работы по замене агрегатов')
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

