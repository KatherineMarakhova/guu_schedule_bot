import config as cf
import telebot
from telebot import types
from Direct import *


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
        i = 0
        for name in my_direct.list_insts:
            btn = types.InlineKeyboardButton(text=str(name), callback_data=f'{i}inst')
            i+=1
            markup.add(btn)

        # bot.send_message(call.message.chat.id, f"Бакалавриат {course}-курс. Институты ГУУ:", reply_markup = markup)
        bot.edit_message_text(f"Бакалавриат {course}-курс. Институты ГУУ:", chat_id = call.message.chat.id, message_id = msg.message_id, reply_markup=markup)

    # получили название института
    if req[0][1:] == 'inst':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        markup = types.InlineKeyboardMarkup()

        my_direct.set_inst(my_direct.list_insts[int(req[0][:1])])      # добавили в наш объект выбранный институт
        my_direct.get_list_napr()                                 # формируем лист направлений

        i = 0
        for name in my_direct.list_napr:
            if str(name) == 'None':continue

            btn = types.InlineKeyboardButton(text=name, callback_data=f'{i}napr')
            i+=1
            markup.add(btn)

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

    # получили название обр программы
    if req[0][1:] == 'edup':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        my_direct.set_edup(my_direct.list_edup[int(req[0][:1])])

        my_direct.get_list_group()  # формируем лист групп

        markup = types.InlineKeyboardMarkup()

        i = 0
        for group in my_direct.list_groups:
            if str(group) == 'None': continue

            btn = types.InlineKeyboardButton(text = group, callback_data = f'{i}group')
            i+=1
            markup.add(btn)
        bot.send_message(call.message.chat.id, "Отлично! Теперь нужно выбрать группу", reply_markup = markup)

    print(f'req[0]: {req[0]}')
    print(f'req[0]: {req[0]=="group"}')
    # получили группу
    if req[0][1:] == 'group':
        # bot.delete_message(call.message.chat.id, call.message.message_id)
        my_direct.set_group(my_direct.list_groups[int(req[0][:1])])

        bot.send_message(call.message.chat.id, "Настройка завершена!✅\n"
                                               f"Курс: {my_direct.course} \n"
                                               f"Направление подготовки: {my_direct.napr} \n"
                                               f"Образовательная программа: {my_direct.edup} \n"
                                               f"Группа: {my_direct.group} \n")

        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton('расписание целиком', callback_data='fullscd')
        btn2 = types.InlineKeyboardButton('расписание четной недели', callback_data='evenscd')
        btn3 = types.InlineKeyboardButton('расписание нечетной недели', callback_data='oddscd')
        btn4 = types.InlineKeyboardButton('расписание по дням', callback_data='weekdayscd')
        markup.add(btn1, btn2, btn3, btn4)

        bot.send_message(call.message.chat.id, "Выбери формат вывода расписания: \n", reply_markup = markup)

    if str(req[0]) == 'fullscd':
        answer = my_direct.get_full_scd()
        bot.send_message(call.message.chat.id, answer)
    if req[0] == 'evenscd':
        pass
    if req[0] == 'oddscd':
        pass
    if req[0] == 'weekdayscd':
        pass

bot.infinity_polling()
