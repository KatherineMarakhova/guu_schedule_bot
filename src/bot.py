import config as cf
import telebot
from telebot import types
from Direct import *
import time

bot = telebot.TeleBot(cf.token)

my_direct = Direct()              #создаем объект нашего класса


@bot.message_handler(commands=['start'])
def button_message(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)

    item1 = types.KeyboardButton("Выбрать курс")
    markup.add(item1)

    bot.send_message(message.chat.id,'Привет! \nЯ бот-хранитель твоего распиcания!', reply_markup = markup)

@bot.message_handler(commands=['full'])
def button_message(message):
    answer = my_direct.get_full_scd()
    bot.send_message(message.chat.id, answer)

@bot.message_handler(content_types=['text'])
def message_reply(message):
    if message.text == "Выбрать курс":
        markup = types.InlineKeyboardMarkup()

        for i in range(1, 5):
            btn = types.InlineKeyboardButton(text = f'{i}-курс', callback_data = f'{i}-курс')
            markup.add(btn)
        bot.send_message(message.chat.id, "Бакалавриат. Курсы: ", reply_markup = markup)


@bot.callback_query_handler(func=lambda call:True)
def callback_query(call):
    req = call.data.split('_')

    # получили курс
    if req[0].find('курс') and my_direct.course == '':
        bot.send_message(call.message.chat.id, "Я не сломался, просто гружу и обрабатываю твое расписание :)\nНужно немного подождать")
        course = str(req[0])
        course = int(course[0])
        my_direct.set_course(course)
        time.sleep(1)
        if my_direct.first_start:
            my_direct.first_start()
            time.sleep(2)

        markup = types.InlineKeyboardMarkup()

        my_direct.get_list_inst()

        for name in my_direct.list_insts:
            btn = types.InlineKeyboardButton(text=str(name), callback_data=str(name))
            markup.add(btn)
        bot.send_message(call.message.chat.id, f"Бакалавриат {course}-курс. Институты ГУУ:", reply_markup=markup)

    # получили название института
    if req[0] in my_direct.list_insts:

        markup = types.InlineKeyboardMarkup()

        my_direct.set_inst(req[0])      # добавили в наш объект выбранный институт
        my_direct.get_dict_napr()       # формируем лист направлений

        for name in my_direct.list_napr:
            if str(name) == 'None':continue
            if len(name) > 32:          # колбек не поддерживает текст больше 32 символов, поэтому его нужно обрезать
                btn = types.InlineKeyboardButton(text = name, callback_data = name[:33])
                markup.add(btn)
                continue
            btn = types.InlineKeyboardButton(text=name, callback_data=name)
            markup.add(btn)
        m = "Направления " + req[0] + ":"
        bot.send_message(call.message.chat.id, m, reply_markup=markup)

    # получили название направления
    if my_direct.check_name_in_list(req[0], my_direct.list_napr): #вторая проверка для укороченных названий
        my_direct.set_napr(req[0])
        my_direct.get_list_group()  # формируем лист групп

        markup = types.InlineKeyboardMarkup()

        print(f'list_groups {my_direct.list_groups}')
        for group in my_direct.list_groups:
            if str(group) == 'None': continue
            btn = types.InlineKeyboardButton(text = group, callback_data = int(group))
            markup.add(btn)
        bot.send_message(call.message.chat.id, "Отлично! Теперь нужно выбрать группу", reply_markup = markup)
    print(f'req[0]: {req[0]}, type: {type(req[0])}')
    if req[0] in str(my_direct.list_groups):

        my_direct.set_group(req[0])
        bot.send_message(call.message.chat.id, "Я умею выводить расписание в разном виде:\n"
                                               "- расписание целиком /full \n"
                                               "- расписание четной недели \n"
                                               "- расписание нечетной недели \n"
                                               "- расписание по дням недели \n")

bot.infinity_polling()
