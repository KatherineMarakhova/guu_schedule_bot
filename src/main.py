from openpyxl import *
from openpyxl.utils import range_boundaries
from pathlib import Path

# Получение пути интересующего нас файла
def get_file_path():
    with Path(r"/Users/katherine.marakhova/PycharmProjects/exampleBot/files") as direction:
        for f in direction.glob("4-курс-бакалавриат*.xlsx"):
            #print(f'ya nashelsya: {f}')
            return f

# Обработка объединенных ячеек, создание
def unmerge_all_cells(path):
    workbook = load_workbook(path)
    n = len(workbook.sheetnames)
    for i in range(n):
        sheet = workbook.worksheets[i]
        set_merges = sorted(sheet.merged_cells.ranges.copy())
        for cell_group in set_merges:
            min_col, min_row, max_col, max_row = range_boundaries(str(cell_group))
            top_left_cell_value = sheet.cell(row=min_row, column=min_col).value
            sheet.unmerge_cells(str(cell_group))
            for row in sheet.iter_rows(min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row):
                for cell in row:
                    cell.value = top_left_cell_value
        workbook.save(path)
    exit()


# ФУНКЦИЯ ВЫДАЕТ ИНДЕКС ЭЛЕМЕНТА ОТНОСИТЕЛЬНО ТАБЛИЦЫ
def get_indexes(sheet, header_el):
     for i in range(1, sheet.max_row):
          for j in range(sheet.max_column):
               if (sheet[i][j].value == header_el):
                    #print(header_el, ' находится на (i, j): ', i , j, '\n')
                    return (i, j)

"""
        Функция возвращает словарь наименований относительно СТРАНИЦЫ
        Название в категории(Институт/Направление/Образовательная программа) - индекс по строке, индекс по столбцу
    """
def get_indexes_category_sheet(sheet, category_name):
    start_indexes = get_indexes(sheet, category_name)
    y = start_indexes[0]                                             #индекс строки
    x = start_indexes[1]                                             #индекс столбца
    dict_of = {}
    temp = ''
    for j in range(x+3, sheet.max_column):                      # сдвигаемся вправо на три элемента
          name = str(sheet[y][j].value).strip()
          if (name != '' and name != temp):
               # if name == "НАПРАВЛЕНИЕ" or name == "ИНСТИТУТ": continue
               #print('добавил:', name, '\nиндекс:', [y, j])
               dict_of[name] = [y, j]
          else:
               continue
          temp = name                                                 #сохраняем значение для проверки на объединенную ячейку
    return dict_of

# Функция возвращает список элементов по любой категории относительно ВСЕГО файла(Институт/Направление/Образовательная программа)
def get_list_category_file(wb, category_name):
    n = len(wb.sheetnames)
    list_cat = []
    for i in range(n):
        sheet = wb.worksheets[i]
        dict_inst = get_indexes_category_sheet(sheet, category_name)
        list_cat.append(list(dict_inst.keys()))
    return list_cat



path = get_file_path()
first_open = False
if (first_open):                # Если открываем первый раз, то обрабатываем таблицу и перезаписываем
    unmerge_all_cells(path)
wb = load_workbook(path)

sheet = wb.active
sheet = wb.worksheets[0]


# full_inst_list = get_list_category_file(wb, 'ИНСТИТУТ')
full_inst_list = wb.sheetnames

n = len(full_inst_list)

for i in range(n):                     # так надо
    s = full_inst_list[i]
    if s.find(',') != -1:
        pieces = s.split(',')
        full_inst_list.remove(s)
        for i in pieces:
            if i.find('4 курс')==-1:
                i += ' 4 курс'
            full_inst_list.append(i)
    if s.startswith('3'):
        full_inst_list.remove(s)

print(full_inst_list)
# full_edup_list = get_list_category_file(wb, 'Образовательная программа')

# print(get_indexes(sheet, "НАПРАВЛЕНИЕ"))
#print(get_indexes_category_sheet(sheet, get_indexes(sheet, 'НАПРАВЛЕНИЕ')))
# print(get_list_category_file(wb, 'ИНСТИТУТ'))
# list_inst = get_list_category_file(wb, 'НАПРАВЛЕНИЕ')
# print(list_inst)



# idx = get_indexes(sheet, 'ИНСТИТУТ')
# # print(idx)
# names_inst = get_indexes_category_sheet(sheet, idx)
# print(names_inst)
#
# idx2 = get_indexes(sheet, 'НАПРАВЛЕНИЕ')
# #print(idx2)
# dict_napr = get_indexes_of_category(sheet, idx2)
# #print(dict_napr)
# next_napr_idx = get_indexes_of_next_napr(dict_napr, 'ГОСУДАРСТВЕННОЕ И МУНИЦИПАЛЬНОЕ УПРАВЛЕНИЕ')
# #print(f'Индекс след. направления: {next_napr_idx}')
# list_naprs = get_full_scd_by_napr(sheet, dict_napr, 'ГОСУДАРСТВЕННОЕ И МУНИЦИПАЛЬНОЕ УПРАВЛЕНИЕ')
# #print(list_naprs)
#
# idx3 = get_indexes(sheet, 'Образовательная программа')
# # print(idx3)
# names_edup = get_indexes_of_edu_program(sheet, idx3)
# # print(names_edup)
