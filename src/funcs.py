
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
