import config as cf
import telebot
from telebot import types
from Direct import *
import time

bot = telebot.TeleBot(cf.token)

my_direct = Direct()              #создаем объект нашего класса


@bot.message_handler(commands=['start'])
def button_message(message):
    # markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup = types.InlineKeyboardMarkup()
    btn = types.InlineKeyboardButton("Выбрать курс", callback_data = 'start')
    markup.add(btn)

    bot.send_message(message.chat.id,'Привет! \nЯ бот-хранитель твоего распиcания!', reply_markup = markup)

@bot.message_handler(commands=['full'])
def button_message(message):
    answer = my_direct.get_full_scd()
    bot.send_message(message.chat.id, answer)

# @bot.message_handler(content_types=['text'])
# def message_reply(message):
#     if message.text == "Выбрать курс":
#         markup = types.InlineKeyboardMarkup()
#
#         for i in range(1, 5):
#             btn = types.InlineKeyboardButton(text = f'{i}-курс', callback_data = f'{i}-курс')
#             markup.add(btn)
#         bot.send_message(message.chat.id, "Бакалавриат. Курсы: ", reply_markup = markup)


@bot.callback_query_handler(func=lambda call:True)
def callback_query(call):
    req = call.data.split('_')

    if req[0] == 'start':
        # bot.delete_message(call.message.chat.id, call.message.message_id)
        bot.edit_message_text('Привет! \nЯ бот-хранитель твоего распиcания!', chat_id=call.message.chat.id, message_id=call.message.message_id)
        markup = types.InlineKeyboardMarkup()

        for i in range(1, 5):
            btn = types.InlineKeyboardButton(text=f'{i}-курс', callback_data=f'{i}course')
            markup.add(btn)
        bot.send_message(call.message.chat.id, "Бакалавриат. Курсы: ", reply_markup=markup)


    # получили курс
    if req[0][1:] == 'course':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        msg = bot.send_message(call.message.chat.id, "Загрузка..")
        course = int(str(req[0])[0])
        if course != my_direct.course:
            my_direct.clean_all()
            my_direct.set_course(course)
            my_direct.first_start()

        markup = types.InlineKeyboardMarkup()

        my_direct.get_list_inst()

        for name in my_direct.list_insts:
            btn = types.InlineKeyboardButton(text=str(name), callback_data=str(name))
            markup.add(btn)

        # bot.send_message(call.message.chat.id, f"Бакалавриат {course}-курс. Институты ГУУ:", reply_markup = markup)
        bot.edit_message_text(f"Бакалавриат {course}-курс. Институты ГУУ:", chat_id = call.message.chat.id, message_id = msg.message_id, reply_markup=markup)


    # получили название института
    if req[0] in my_direct.list_insts:
        bot.delete_message(call.message.chat.id, call.message.message_id)
        markup = types.InlineKeyboardMarkup()

        my_direct.set_inst(req[0])      # добавили в наш объект выбранный институт
        my_direct.get_list_napr()       # формируем лист направлений
        i = 0
        for name in my_direct.list_napr:
            if str(name) == 'None':continue

            btn = types.InlineKeyboardButton(text=name, callback_data=f'{i}napr')
            markup.add(btn)
        m = "Направления " + req[0] + ":"
        bot.send_message(call.message.chat.id, f"Направления {my_direct.inst}:", reply_markup=markup)

    # получили название направления
    if req[0][1:] == 'napr':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        my_direct.set_napr(my_direct.list_napr[int(req[0][:1])])

        my_direct.get_list_edup() # формируем лист обр программ

        markup = types.InlineKeyboardMarkup()
        print(f'my_direct.list_edup: {my_direct.list_edup}')
        i = 0

        for edup in my_direct.list_edup:
            edup = str(edup)
            if edup == 'None': continue
            btn = types.InlineKeyboardButton(text = edup, callback_data = f'{i}edup')
            i+=1
            markup.add(btn)
        bot.send_message(call.message.chat.id, f"Образовательные программы {my_direct.napr}:", reply_markup = markup)

    print(f'req[0]: {req[0]}, type: {type(req[0])}')

    # получили название обр программы
    if req[0][1:] == 'edup':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        print(f'my_direct.list_edup[0]: {my_direct.list_edup[0]}')
        my_direct.set_edup(my_direct.list_edup[int(req[0][:1])])
        my_direct.get_list_group()  # формируем лист групп

        markup = types.InlineKeyboardMarkup()

        print(f'list_groups {my_direct.list_groups}')
        for group in my_direct.list_groups:
            if str(group) == 'None': continue
            btn = types.InlineKeyboardButton(text = group, callback_data = int(group))
            markup.add(btn)
        bot.send_message(call.message.chat.id, "Отлично! Теперь нужно выбрать группу", reply_markup = markup)

    print(f'req[0]: {req[0]}, type: {type(req[0])}')

    # получили группу
    if req[0] in str(my_direct.list_groups):
        bot.delete_message(call.message.chat.id, call.message.message_id)
        my_direct.set_group(req[0])
        bot.send_message(call.message.chat.id, "Настройка завершена!\n"
                                               f"{my_direct.napr} {my_direct.edup} {my_direct.group}\n"
                                               "Выбери способ вывода расписания:\n"
                                               "- расписание целиком /full \n"
                                               "- расписание четной недели \n"
                                               "- расписание нечетной недели \n"
                                               "- расписание по дням недели \n")

bot.infinity_polling()
