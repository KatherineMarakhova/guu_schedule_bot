from openpyxl import *
from openpyxl.utils import range_boundaries
from pathlib import Path
import selenium_fcs as sf

class Direct:
    course = ''
    path = ''
    wb = ''
    sheet = ''
    list_insts = []         # список институтов
    inst = ''               # выбранный институт
    list_edup = []
    edup = ''
    list_napr = []
    napr = ''
    list_groups = []
    group = ''
    msg_count = 0   # счетчик сообщений
    check_msg = ''  # тут хранится id сообщения, которое надо изменить при выводе расписания по дням


    def set_path(self, path):
        self.path = path

    def set_course(self, course):
        self.course = course

    def set_inst(self, inst):
        self.inst = inst
        self.sheet = self.wb[inst]

    def set_napr(self, napr):
        self.napr = napr

    def set_edup(self, edup):
        self.edup = edup

    def set_group(self, group):
        self.group = group

    def get_list_inst(self):
        full_inst_list = self.wb.sheetnames
        # Иногда попадаются файлы с лишними страницами, тут их удаляем
        for s in full_inst_list:
            if s.startswith('3') or s == 'None': # если строка начинается с "3". в файле с расписанием 3 курса строки не начинаются с 3
                full_inst_list.remove(s)
        self.list_insts = full_inst_list


    # Получение словаря {'название инст/напр': его индекс относительно талицы(строка, столбец)}
    def get_list_napr(self):
        # 1. определяем строку с направлениями
        # 2. определяем столбец с которого начинаются наименования
        curr = (self.next_idx_cat(self.sheet, 'НАПРАВЛЕНИЕ', 'НАПРАВЛЕНИЕ'))      # строка и столбец

        temp = ''
        list_napr = []

        for j in range(curr[1], self.sheet.max_column+1):                           # сдвигаемся вправо на два элемента(ЭТО НАДО ИСПРАВИТЬ)
              name = str(self.sheet.cell(curr[0],j).value).strip()
              if (name != '' and name != temp and name!='None'):
                   list_napr.append(name)
              else:
                   continue
              temp = name
        self.list_napr = list_napr

    # получаем образовательные программы относительно направления
    def get_list_edup(self):
        row = self.get_indexes(self.sheet, 'Образовательная программа')[0]            # рабочая строка
        coll = self.get_indexes_cat(self.sheet, self.napr, 'направление')[1]            # столбец где начинается направление
        next = self.next_idx_cat(self.sheet, self.napr, 'направление')[1]             # столбец где оно заканчивается
        list_edup = []
        temp = ''
        for j in range(coll, next):  # идем по столбцам
            name = str(self.sheet.cell(row, j).value).strip()
            if (name != '' and name != temp and name != 'None'):
                list_edup.append(name)
            else:
                continue
            temp = name  # сохраняем значение для проверки на объединенную ячейку

        self.list_edup = list_edup

    def get_list_group(self):
        row = (self.get_indexes(self.sheet, '№ учебной группы'))[0]                           # рабочая строка
        coll = (self.get_indexes_cat(self.sheet, self.edup, 'Образовательная программа'))[1]  # столбец где начинается обр прог
        next = (self.next_idx_cat(self.sheet, self.edup, 'Образовательная программа'))[1]     # столбец где заканчивается обр прог

        groups = []
        for j in range(coll, next):
            group = self.sheet.cell(row, j).value
            if group == "None": continue
            groups.append(group)
        self.list_groups = groups

    def check_name_in_list(self, name, somelist):
        for i in somelist:
            if i.find(name) != -1:
                return True
        return False

    #  БЛОК ПОЛУЧЕНИЯ И ОБРАБОТКИ ФАЙЛА ==========================================================
    def first_start(self):

        self.path = self.get_file_path()             # записываем путь скачанного файла
        self.unmerge_all_cells()                     # обрабатываем объединенные ячейки
        self.unmerge_institutes()                    # разделяем два института, хранящихся на одном листе
        # print(f'path: {self.path}')
        self.wb = load_workbook(self.path)

    def clean_all(self):
        self.path = ''
        self.wb = ''
        self.sheet = ''
        self.list_insts = ''  # список институтов
        self.inst = ''  # выбранный институт
        self.list_edup = ''
        self.edup = ''
        self.list_napr = ''
        self.napr = ''
        self.list_groups = ''
        self.group = ''

    # Получение пути интересующего нас файла
    def get_file_path(self):
        with Path(r"../files") as direction:
            s = str(self.course) + "-курс-бакалавриат*.xlsx"
            path = ''
            for f in direction.glob(s):
                path = f
                break
            if path:
                return path
            else:
                sf.get_file(self.course)  # скачиваем файл с сайта относительно курса
                for f in direction.glob(s):
                    return f

    # Обработка объединенных ячеек, создание
    def unmerge_all_cells(self):
        workbook = load_workbook(self.path)
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
            workbook.save(self.path)

    # Разделение институтов, хранящихся на одном листе
    def unmerge_institutes(self):
        workbook = load_workbook(self.path)
        for s in workbook.sheetnames:
            if s.find(',') != -1:  # название нашего листа 'ИУПСиБК, ИИС 4 курс'
                sheet1 = workbook[s]

                t = s.find('курс') - 2  # вычленяем курс из названия листа
                z = s.find(',')  # находим позицию запятой
                fname = s[:z] + ' ' + s[t:]  # иупсибк
                sname = s[z + 2:]  # иис
                # Теперь лист на котором мы были будет называться как sname, а новый как fname
                sheet1.title = sname  # переименовыем в ИИС 4 курс
                workbook.create_sheet(fname)  # создаем лист ИУПСиБК 4 курс
                workbook.save(self.path)  # страхуемся, сохраняем наш док

                # переопределяем листы, так четче видно что где
                sheet1 = workbook[sname]  # иис тут заполнено
                sheet2 = workbook[fname]  # иупсибк тут пусто

                last_inst_name = ''
                y, x = self.get_indexes(sheet1, 'ИНСТИТУТ')
                for i in range(1, sheet1.max_column):
                    # val = sheet1[y][i].value
                    val = sheet1.cell(y, i).value
                    if val != 'None':
                        last_inst_name = val

                inst_idx = self.get_indexes(sheet1, last_inst_name)  # индекс с которого начинается второй институт(ИИС)

                # print(f'Последний институт {last_inst_name}, находится {inst_idx}')

                # заполняем новый лист(ИУПСиБК)
                for i in range(1, sheet1.max_row):
                    for j in range(1, inst_idx[1]):
                        sheet2.cell(i, j).value = sheet1.cell(i, j).value

                # print(f'inst_idx[1]: {inst_idx[1]}')
                # удлаляем ненужные столбцы с исходного листа
                sheet1.delete_cols(idx=5, amount=(inst_idx[1] - 5))  # тут пока костыль в виде 4 - именно столько столбцов нужно отступить слева
                # надо будет написать функцию добывающую этот индекс, чтобы было гибко
                workbook.save(self.path)

    # ✔️️ Получение индекса(строка, столбец). Используется для значений Институт, направление, образовательная программа и тд.(категории)
    def get_indexes(self, sheet, category):
        category = (category.strip()).lower()
        for i in range(1, sheet.max_row):
            for j in range(1, sheet.max_column+1):
                val = (str(sheet.cell(i, j).value).strip()).lower()
                if (val == category):
                    return (i, j)

    # Получение индекса названия относительно категории(института, направления и тд.)
    def get_indexes_cat(self, sheet, name, category):
        name = (name.strip()).lower()
        y, x = self.get_indexes(self.sheet, category)    # нам нужна только строка! те у
        for i in range(y, sheet.max_row):
            for j in range(1, sheet.max_column+1):
                val = str(sheet.cell(i, j).value).strip().lower()

                if (val == name):
                    return (i, j)

        # Получение индекса другого элемента по строке(не равного по названию)
    def next_idx_cat(self, sheet, name, category):
        row = self.get_indexes(self.sheet, category)[0]      #строка
        name = name.strip().lower()
        coll = self.get_indexes(sheet, name)[1]              #столбец

        if coll == sheet.max_column: return (row, coll+1)

        for c in range(coll, sheet.max_column+1):  # строка не меняется
            val = str(sheet.cell(row, c).value).strip().lower()
            if val == name:
                if c == sheet.max_column:
                    return (row, sheet.max_column + 1)            #если это и есть наш последний элемент, то возращем его + 1
                continue
            elif val != name: return (row, c)
        return

    # БЛОК ВЫВОДА РАСПИСАНИЯ ======================================================================
    def get_scd_full(self):
        answer = f'Расписание для {self.edup.capitalize()} {self.course}-{self.group}\n'
        # получаем граничные индексы
        day = self.get_indexes(self.sheet, 'Понедельник')
        d_idx = int(day[1])
        row = day[0]                                                                        # строка
        coll = self.get_indexes_cat(self.sheet, self.edup, 'Образовательная программа')[1]  # столбец
        coll = coll + int(self.group) - 1                                                   # столбец относительно группы

        temp = ''
        lesson = ''
        for i in range(row, self.sheet.max_row):

            day = str(self.sheet.cell(i, d_idx).value)
            time = str(self.sheet.cell(i, d_idx+1).value)
            even = str(self.sheet.cell(i, d_idx+2).value)
            sbj = str(self.sheet.cell(i, coll).value)

            if (day.lower() == 'none'): break

            if sbj == 'None': sbj = 'Занятий нет'

            dash = '--------------------------------------️'

            if temp != day:
                spaces = ''
                n = len(dash) - len('❗️' + day + '❗️') - 1
                for i in range(n):
                    spaces += ' '
                answer += f'{dash}\n❗️{day}❗{spaces}|️\n{dash}\n'
                lesson = 1
                answer += (f'{lesson}📍{time}\n')
                temp = day
                lesson += 1

            if i % 2 != 0:
                answer += (f'{lesson}📍{time}\n- {even.capitalize()}\n {sbj}\n\n')
                lesson += 1
            else:
                answer += (f'- {even.capitalize()}\n {sbj}\n\n')

        return answer

    def get_scd_even(self, eveness = "ЧЁТ."):

        answer = f'Расписание для {self.edup} {self.course}-{self.group}\n'
        answer += f'{eveness.capitalize()} неделя\n'

        # получаем граничные индексы
        day = self.get_indexes(self.sheet, 'Понедельник')
        d_idx = int(day[1])
        row = day[0]  # строка
        coll = self.get_indexes_cat(self.sheet, self.edup, 'Образовательная программа')[1]  # столбец
        coll = coll + int(self.group) - 1  # столбец относительно группы


        temp = ''
        lesson = ''
        for i in range(row, self.sheet.max_row):

            day = str(self.sheet.cell(i, d_idx).value)
            if (day.lower() == 'none'): break

            time = str(self.sheet.cell(i, d_idx+1).value)
            even = str(self.sheet.cell(i, d_idx+2).value)
            sbj = str(self.sheet.cell(i, coll).value)

            if sbj == 'None': sbj = 'Занятий нет'

            dash = '--------------------------------------️'

            if temp != day:
                spaces = ''
                n = len(dash) - len('❗️' + day + '❗️') - 1
                for i in range(n):
                    spaces += ' '
                answer += f'{dash}\n❗️{day}❗{spaces}|️\n{dash}\n'
                lesson = 1
                temp = day

            if even.strip().lower() == eveness.strip().lower():
                answer += (f'{lesson}📍{time}\n{sbj}\n\n')
                lesson += 1
        return answer

    def get_scd_weekday(self, weekday = 'Понедельник'):
        answer = f'Расписание для {self.edup} {self.course}-{self.group}\n'

        spaces = ''
        dash = '--------------------------------------️'
        n = len(dash) - len('❗️' + weekday.upper() + '❗️') - 1
        for i in range(n):
            spaces += ' '
        answer += f'{dash}\n❗️{weekday.upper()}❗{spaces}|️\n{dash}\n'

        # получаем граничные индексы
        day = self.get_indexes(self.sheet, 'Понедельник')
        d_idx = int(day[1])
        row = day[0]  # строка
        coll = self.get_indexes_cat(self.sheet, self.edup, 'Образовательная программа')[1]  # столбец
        coll = coll + int(self.group) - 1  # столбец относительно группы

        temp = ''
        lesson = 1
        for i in range(row, self.sheet.max_row):

            day = str(self.sheet.cell(i, d_idx).value)
            if (day.lower() == 'none'): break

            time = str(self.sheet.cell(i, d_idx+1).value)
            even = str(self.sheet.cell(i, d_idx+2).value)
            sbj = str(self.sheet.cell(i, coll).value)

            if sbj == 'None': sbj = 'Занятий нет'
            if weekday.lower() == day.lower():
                if i % 2 != 0:
                    answer += (f'{lesson}📍{time}\n- {even.capitalize()}\n {sbj}\n\n')
                    lesson += 1
                else:
                    answer += (f'- {even.capitalize()}\n {sbj}\n\n')
        return answer

    def clear_attributes(self):
        self.course = ''
        self.path = ''
        self.wb = ''
        self.sheet = ''
        self.list_insts = []  # список институтов
        self.inst = ''  # выбранный институт
        self.list_edup = []
        self.edup = ''
        self.list_napr = []
        self.napr = ''
        self.list_groups = []
        self.group = ''
