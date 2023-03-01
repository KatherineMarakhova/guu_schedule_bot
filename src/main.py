from openpyxl import *
from openpyxl.utils import range_boundaries
from pathlib import Path
import selenium_fcs as sf
import time


# Получение пути интересующего нас файла
def get_file_path():
    with Path(r"/Users/katherine.marakhova/PycharmProjects/exampleBot/files") as direction:
        for f in direction.glob("3-курс-бакалавриат*.xlsx"):
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

# Разделение институтов, хранящихся на одном листе
def unmerge_institutes(path):
    workbook = load_workbook(path)
    for s in workbook.sheetnames:
        if s.find(',') != -1:                           #название нашего листа 'ИУПСиБК, ИИС 4 курс'
            sheet1 = workbook[s]

            t = s.find('курс')-2                        #вычленяем курс из названия листа
            z = s.find(',')                             #находим позицию запятой
            fname = s[:z] + ' ' + s[t:]                 #иупсибк
            sname = s[z+2:]                             #иис
            # Теперь лист на котором мы были будет называться как sname, а новый как fname
            sheet1.title = sname                        # переименовыем в ИИС 4 курс
            workbook.create_sheet(fname)                # создаем лист ИУПСиБК 4 курс
            workbook.save(path)                         # страхуемся, сохраняем наш док

            #переопределяем листы, так четче видно что где
            sheet1 = workbook[sname]                    # иис тут заполнено
            sheet2 = workbook[fname]                    # иупсибк тут пусто

            inst_idx = next_idx(sheet1, 'ИНСТИТУТ')     # индекс с которого начинается второй институт(ИИС)

            # заполняем новый лист(ИУПСиБК)
            for i in range(1, sheet1.max_row):
                for j in range(1, inst_idx[1]):
                    sheet2.cell(i, j).value = sheet1.cell(i, j).value

            # удлаляем ненужные столбцы с исходного листа
            sheet1.delete_cols(idx=4, amount=(inst_idx[1]-4)) # тут пока костыль в виде 4 - именно столько столбцов нужно отступить слева
                                                              # надо будет написать функцию добывающую этот индекс, чтобы было гибко
            workbook.save(path)

# Получение индекса последнего имени института/чего-то еще другого(для разъединения двух институтов на одном листе)
def next_idx(sheet, category):
        y, x = get_indexes(sheet, category) # строка, столбец
        temp = sheet.cell(y, x).value
        new_x = 0
        # строка не меняется
        for j in range(x, sheet.max_column):
            if (sheet.cell(y,j).value != temp and sheet.cell(y,j).value != 'None'):
                temp = sheet.cell(y,j).value
                new_x = j
        return (y, new_x)
# Получение индекса(строка, столбец) относительно ячейки(для определения строки(столбца) с назв. институтов/направлений и др.
def get_indexes(sheet, header_el):
         for i in range(1, sheet.max_row):
              for j in range(sheet.max_column):
                   if (sheet[i][j].value == header_el):
                        #print(header_el, ' находится на (i, j): ', i , j, '\n')
                        return (i, j)


class Direct:
    path = ''
    wb = ''
    list_insts = '' #список институтов
    inst = ''  #выбранный институт
    list_napr = ''
    dict_napr = ''
    napr = ''

    def set_path(self, path):
        self.path = path
        self.wb = load_workbook(path)

    def set_inst(self, inst):
        self.inst = inst

    def set_napr(self, napr):
        self.napr = napr

    def get_list_inst(self):
        full_inst_list = self.wb.sheetnames
        # Иногда попадаются файлы с лишними страницами, тут их удаляем
        for s in full_inst_list:
            if s.startswith('3') or s == 'None':
                full_inst_list.remove(s)
        self.list_insts = full_inst_list


    # Получение индекса последнего имени института/чего-то еще другого(для разъединения двух институтов на одном листе)
    def next_idx(self, sheet, category):
        y, x = get_indexes(sheet, category) # строка, столбец
        temp = sheet.cell(y, x).value
        new_x = 0
        # строка не меняется
        for j in range(x, sheet.max_column):
            if (sheet.cell(y,j).value != temp and sheet.cell(y,j).value != 'None'):
                temp = sheet.cell(y,j).value
                new_x = j
        return (y, new_x)

    # Получение индекса(строка, столбец) относительно ячейки(для определения строки(столбца) с назв. институтов/направлений и др.
    def get_indexes(self, sheet, header_el):
         for i in range(1, sheet.max_row):
              for j in range(sheet.max_column):
                   if (sheet[i][j].value == header_el):
                        #print(header_el, ' находится на (i, j): ', i , j, '\n')
                        return (i, j)

    # Получение словаря {'название инст/напр': его индекс относительно талицы(строка, столбец)}
    def get_dict_napr(self):
        sheet = self.wb[self.inst]
        start_indexes = get_indexes(sheet, 'НАПРАВЛЕНИЕ')
        y = start_indexes[0]                                             #индекс строки
        x = start_indexes[1]                                             #индекс столбца
        dict_of = {}
        temp = ''
        for j in range(x+3, sheet.max_column):                      # сдвигаемся вправо на три элемента
              name = str(sheet[y][j].value).strip()
              if (name != '' and name != temp):
                   dict_of[name] = [y, j]
              else:
                   continue
              temp = name                                                 #сохраняем значение для проверки на объединенную ячейку
        self.dict_napr = dict_of
        self.list_napr = list(dict_of.keys())


# Получаем путь к нашему файлу. Пригодится в дальнейшем, когда захотим выыводить разные курсы
path = get_file_path()

"""
1. Скачиваем файл с сайта
2. Вызываем функцию по очистке объединенных ячеек (unmerge_all_cells)
3. Вызываем функцию по разделению институтов, хранящихся на одном листе (unmerge_institutes)
пока усе
"""

first_start = False
def first_start():
    sf.get_file()
    path = get_file_path()
    unmerge_all_cells(path)
    unmerge_institutes(path)
    return path

obj = Direct()
# first_start()
obj.set_path('/Users/katherine.marakhova/PycharmProjects/exampleBot/files/3-курс-бакалавриат-ОФО-50.xlsx')
obj.get_list_inst()
print(obj.list_insts)

obj.set_inst('ИГУиП 3 курс')
print(obj.inst)

obj.get_dict_napr()
print(obj.list_napr)



# # Список институтов
# full_inst_list = wb.sheetnames
# # Иногда попадаются файлы с лишними страницами, тут их удаляем
# for s in full_inst_list:
#     if s.startswith('3') or s == 'None':
#         full_inst_list.remove(s)
# print(full_inst_list)
#
# # get_indexes_category_sheet(wb, inst, 'НАПРАВЛЕНИЕ')

















