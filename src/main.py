from openpyxl import *
from openpyxl.utils import range_boundaries

def unmerge_all_cells(workbook):
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
        workbook.save("openpyxl_merge_unmerge.xlsx")
    exit()

# ФУНКЦИЯ ВЫДАЕТ ИНДЕКС ЭЛЕМЕНТА ОТНОСИТЕЛЬНО ТАБЛИЦЫ
def get_indexes(sheet, header_el):
     for i in range(1, sheet.max_row):
          for j in range(sheet.max_column):
               if (sheet[i][j].value == header_el):
                    #print(header_el, ' находится на (i, j): ', i , j, '\n')
                    return (i, j)

# ФУНКЦИЯ ВЫДАЕТ СЛОВАРЬ {ИНСТИТУТ/НАПРАВЛЕНИЕ : ИНДЕКС_ПО_СТРОКЕ, ИНДЕКС_ПО_СТОЛБЦУ}
def get_indexes_of_category(sheet, start_indexes):
     y = start_indexes[0]                                             #индекс строки
     x = start_indexes[1]                                             #индекс столбца
     dict_of = {}
     temp = ''
     for j in range(x+3, sheet.max_column):                      # сдвигаемся вправо на три элемента
          name = str(sheet[y][j].value).strip()
          if (name == "None"): continue
          if (name != '' and name != temp):
               if name == "НАПРАВЛЕНИЕ" or name == "ИНСТИТУТ": continue
               #print('добавил:', name, '\nиндекс:', [y, j])
               dict_of[name] = [y, j]
          else:
               continue
          temp = name                                                 #сохраняем значение для проверки на объединенную ячейку

     return dict_of

# ФУНКЦИЯ ВЫДАЕТ СЛОВАРЬ {ОБРАЗОВАТЕЛЬНАЯ ПРОГРАММА : ИНДЕКС_ПО_СТРОКЕ, ИНДЕКС_ПО_СТОЛБЦУ}
def get_indexes_of_edu_program(sheet, start_indexes):
     ''' Данная функция предназначена для строки ОБРАЗОВАТЕЛЬНАЯ ПРОГРАММА.
         Особенность в том, что в случае с объединенной ячейкой будем брать значение "сверху",
         т.е. из НАПРАВЛЕНИЯ.
     '''
     y = start_indexes[0]  # индекс строки
     x = start_indexes[1]  # индекс столбца
     dict_of = {}
     temp = '-1'
     k = 0 #счетчик для проверки при вставке верхнего значения
     for j in range(x+3, sheet.max_column-1):                            # тут костыль +3 пусть пока так
          # у обр программы может быть и повторяющееся и пустое
          # в случае с повторяющимся надо записывать первое
          # в случае с пустым верхнее
          name = str(sheet[y][j].value).strip()                   # удаляем лишние пробелы
          if (name != '' and name != temp and name!="None"):
               # print('добавил:', name, '\nиндекс:', [y, j])
               dict_of[name] = [y, j]
               temp = name                                            # сохраняем значение для проверки на объединенную ячейку
          elif(name == temp):
               continue
          elif(name == '' or name == "None"):
               name = str(sheet[y-1][j].value).strip()
               if(name != temp):
                    dict_of[name] = [y, j]
                    temp = name
     return dict_of

# ФУНКЦИЯ ВЫДАЕТ СЛОВАРЬ {СЛЕД_НАПРАВЛЕНИЕ : ИНДЕКС_ПО_СТРОКЕ, ИНДЕКС_ПО_СТОЛБЦУ}
def get_indexes_of_next_napr(dict_of_napr, name_napr):
     # возвращает индекс следующего элемента, а если такового нет, то индекс самого себя
     find = False
     for key in dict_of_napr:
          if key == name_napr:
               find = True
               continue
          if find == True:
               return dict_of_napr[key]
     return dict_of_napr[name_napr]


def get_full_scd_by_napr(sheet, dict_of_napr, name_napr):
    '''
    Чтобы вывести расписание по направлению нам нужно:
    - имя направления
    - граничные индексы
    - строка ответа
    '''

    answer = ''
    # получаем граничные индексы
    y = (dict_of_napr[name_napr])[0]
    start_x = (dict_of_napr[name_napr])[1]

    next_idx = get_indexes_of_next_napr(dict_of_napr, name_napr)
    end_x = next_idx[1]
    if start_x == end_x: end_x = (sheet.max_column) # функция, выдающая следующий элемент, возвращает индекс себя самого, если он стоит в конце

    nrows = sheet.max_row + 1  # количество строк(+1 тк нумерация с единицы)

    gr = 1  # Счетчик групп
    les = 1  # Счетчик пар
    weekday = ""  # ДЕНЬ НЕДЕЛИ

    # ПОКА БУДЕМ ВЫВОДИТЬ РАСПИСАНИЕ ПО КАЖДОМУ НАПРАВЛЕНИЮ ПО ОТДЕЛЬНОСТИ, Т.Е. ДРУГ ЗА ДРУГОМ ПРОПИСЫВАТЬ РАСПИСАНИЕ ПО ГРУППАМ(УЖАС!!!!)
    for j in range(start_x, end_x):
        answer = f"Расписание для {name_napr}4-{gr}\n"
        print(f'Расписание для {name_napr}4 -{gr}')
        gr += 1
        for i in range(8, nrows):  # идем по строкам начиная с 3(тк пока не удалось удалить)

            timing = sheet[i][2].value  # ВРЕМЯ ПАРЫ (если есть)
            even_week = sheet[i][3].value  # ЧЕТНОСТЬ НЕДЕЛИ

            subject = sheet[i][4].value  # НАЛИЧИЕ ЗАНЯТИЙ

            if (sheet[i][1].value != "" and sheet[i][1].value != "None"):  # ДЕНЬ НЕДЕЛИ
                les = 1  # возвращаем счетчик к 1
                answer += "---------------------------\n"
                print("---------------------------")
                weekday = sheet[i][1].value
                answer += f'{weekday}\n'
                print(weekday)

            if (timing != "" and timing != "None"):  # ВРЕМЯ (если есть то пишем)
                answer += "-----\n"
                print('-----')
                answer += f'{les} пара {timing}\n'
                print(les, 'пара ', timing)
                les += 1

            answer += f'{even_week}\n'
            print(even_week)  # ЧЕТНОСТЬ НЕДЕЛИ

            if (subject != "" and subject != "None"):  # НАЛИЧИЕ ЗАНЯТИЙ
                answer += f'\t{subject}\n'
                print('\t', subject)
            else:
                answer += f'\tНет занятий'
                print('\t', "Нет занятий")

            if (weekday == "СУББОТА" and even_week == "ЧЁТ."):
                answer += '==================================================================================\n'
                print("==================================================================================")
                answer += 'Конец расписания.\n'
                print("Конец расписания.")
                break
    return answer

def get_scd_napr(sheet, dict_napr, name_napr):
    # первое - получение граничных значений по направлению
    y = (dict_napr[name_napr])[0] # строка
    x = (dict_napr[name_napr])[1] # столбец
    next_idxs = get_indexes_of_next_napr(dict_napr, name_napr)
    end_x = next_idxs[1]

    gr = 1  # Счетчик групп
    num_lesson = 1  # Счетчик пар

    if x == end_x:                  # функция, выдающая следующий элемент, возвращает индекс себя самого,
        end_x = (sheet.max_column)  # если он стоит в конце

    # идем по столбцам! записывая строки
    for j in range(x, end_x):
        for i in range(y, sheet.max_row+1):
            week_day = sheet[i][1].value
            time_pair = sheet[i][2].value
            even_week = sheet[i][3].value

            if(week_day != "None"):
                print(f'День недели:{week_day}')
            if(time_pair != "None"):
                print(f'Пара №{num_lesson}. Время занятий: {time_pair}')
            if(even_week != "None"):
                print(f'Четность недели:{even_week}')


wb = load_workbook('../files/raspisanie.xlsx')
sheet = wb.active
sheet = wb.worksheets[4]
unmerge_all_cells(wb)
# print(sheet.merged_cells.ranges)
# set_merges = sorted(sheet.merged_cells.ranges.copy())


#
# for cell_group in set_merges:
# 	min_col, min_row, max_col, max_row = range_boundaries(str(cell_group))
# 	top_left_cell_value = sheet.cell(row=min_row, column=min_col).value
# 	sheet.unmerge_cells(str(cell_group))
# 	for row in sheet.iter_rows(min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row):
# 		for cell in row:
# 			cell.value = top_left_cell_value
# wb.save("openpyxl_merge_unmerge.xlsx")
# exit()
#


# idx = get_indexes(sheet, 'ИНСТИТУТ')
# # print(idx)
# names_inst = get_indexes_of_category(sheet, idx)
# # print(names_inst)
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
