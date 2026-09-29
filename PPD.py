import sys

#sys.path.append("../Харампур/venv/Lib/site-packages") C:\Users\Admin\PSO\venv
sys.path.append("../PSO/venv/Lib/site-packages/openpyxl")
import openpyxl

from openpyxl import load_workbook
#import pandas as pd


# path to file
file_path = r'C:\Users\Admin\PSO\План 09 Н.xlsx'
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
            # print('start_PPD: ',start_PPD, i)
            #print(i)
            #print(ii)
            #print(i.index(ii))
        # найти индекс ППН     
        if ii == 'ППН':
            end_PPD = data.index(i)
            # print('END_PPD: ',i, data.index(i)-1)
            #print(end_PPD)
            start_PPN= data.index(i)
            # print('start_PPN: ',start_PPN, i)
            end_PPN=data.index(data[len(data)-1])
            # print(end_PPN)
            # print('end_PPN: ', data[len(data)-1], data.index(data[len(data)-1]))

# print('start_PPD: ', start_PPD)
row_end=end_PPD-start_PPD
print('row_end: ', row_end)

# ------------------------------------------------------------------------------------------
# Сервис
rows_data = []
rows_data_tr=[]
rows_data_dem=[]
rows_data_mon=[]
rows_data_pnr=[]

for row in ws.iter_rows(min_row=start_PPD+1, max_row=end_PPD, values_only=True):
    # Сервис
    if row[7] == 1:
        rows_data.append(list([1, 'ЦППД-1',row[2],'насос', row[4], f'НА-{row[3]}', row[5], row[1]]))#row[1]
    # Запись данных в справку Сервис
    spravka_PPD=r'C:\Users\Admin\PSO\справка  ЦППД-1.xlsx'
    cols = [1,2,3, 4, 5, 6, 7, 10]
    start_row=14

    # ВАЖНО: без read_only=True, иначе запись невозможна
    wb = load_workbook(filename=spravka_PPD)
    ws = wb.active

    for r_idx, row_data in enumerate(rows_data, start=start_row):
        for col_num, value in zip(cols, row_data):
            ws.cell(row=r_idx, column=col_num).value = value
# ------------------------------------------------------------------------------------
    # Пишем значение в левую верхнюю ячейку (A10)
    #ws['A22'] = 'Объединённая ячейка на 4 строки'

    # Объединяем A10:A13
    #ws.merge_cells('A22:A24')

    #wb.save('merged_column.xlsx')
    # start_A=start_PPD+4+len(rows_data)
    # print('start_A: ',type(start_A), start_A)
    # end_A=start_A+5
    # print('end_A: ', type(end_A), end_A)
    # col=4
    # ws.cell(row=start_A, column=col, value='Объединённая ячейка')
    # ws.merge_cells(start_row=start_A, start_column=col, end_row=end_A, end_column=col)

    # start_row = 22
    # end_row = 24
    # col = 3  # столбец B

    # ws.cell(row=start_row, column=col, value='Объединённая ячейка')
    # ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
# ---------------------------------------------------------------------------------------    
# Запись ППД ТР
    if row[8] == 1:
        rows_data_tr.append(list([1, 'ЦППД-1',row[2],'насос', row[4], f'НА-{row[3]}', row[5], row[1]]))

    spravka_PPD=r'C:\Users\Fomenko.SM\Харампур\September\справка  ЦППД-1.xlsx'
    cols = [1,2,3, 4, 5, 6, 7, 10]
    start_row_tr=34
    
    # ВАЖНО: без read_only=True, иначе запись невозможна
    for r_idx, row_data in enumerate(rows_data_tr, start=start_row_tr):
        for col_num, value in zip(cols, row_data):
            ws.cell(row=r_idx, column=col_num).value = value

    # Запись ППД Демонтаж
    if row[9] == 1:
            rows_data_dem.append(list([1, 'ЦППД-1',row[2], f'НА-{row[3]}',row[5], row[4], row[1]]))
    
    spravka_PPD=r'C:\Users\Fomenko.SM\Харампур\September\справка  ЦППД-1.xlsx'
    cols = [1,2,3, 4, 5, 7, 10]
    start_row_dem=49
    
    # ВАЖНО: без read_only=True, иначе запись невозможна
    for r_idx, row_data in enumerate(rows_data_dem, start=start_row_dem):
        for col_num, value in zip(cols, row_data):
            ws.cell(row=r_idx, column=col_num).value = value

    # Запись ППД Монтаж
    if row[10] == 1:
            rows_data_mon.append(list([1, 'ЦППД-1',row[2],f'НА-{row[3]}',row[5], row[4], row[1]]))
    
    spravka_PPD=r'C:\Users\Admin\PSO\справка  ЦППД-1.xlsx'
    cols = [1,2,3, 4, 6, 7, 10]
    start_row_mon=49+len(rows_data_mon)
    
    # ВАЖНО: без read_only=True, иначе запись невозможна
    for r_idx, row_data in enumerate(rows_data_mon, start=start_row_mon):
        for col_num, value in zip(cols, row_data):
            ws.cell(row=r_idx, column=col_num).value = value


start_A=start_PPD+6+len(rows_data)
print('start_A: ',type(start_A), start_A)
end_A=start_A+3
print('end_A: ', type(end_A), end_A)
col=5
print("COL: ", col)
ws.cell(row=start_A, column=col, value='Объединённая ячейка!!!!!!')
ws.merge_cells(start_row=start_A, start_column=col, end_row=end_A, end_column=col)


wb.save(spravka_PPD)
print("Данные записаны.")
#print('rows_data_tr: ', rows_data_tr)
print('len(rows_data): ',len(rows_data))

