from telebot.types import ReplyKeyboardRemove
import config
import telebot
from telebot import types
from Direct import *

def fullsqd(obj, chatid):
    msg = bot.send_message(chatid, "Загрузка..")
    answer = obj.get_scd_full()
    if len(answer) > 4096:
        bot.edit_message_text(text='Текст расписания слишком большой!\nПожалуйста, выберите другой формат вывода',
                              chat_id = chatid, message_id=msg.message_id)
    else:
        bot.edit_message_text(text=answer, chat_id = chatid, message_id=msg.message_id)

def evenscd(obj, eveness, chatid):
    answer = obj.get_scd_even(eveness)
    bot.send_message(chatid, answer)

def weekdayscd(msg, obj):
    markup = types.InlineKeyboardMarkup(row_width=7)
    btn1 = types.InlineKeyboardButton('ПН', callback_data='Понедельник')
    btn2 = types.InlineKeyboardButton('ВТ', callback_data='Вторник')
    btn3 = types.InlineKeyboardButton('СР', callback_data='Среда')
    btn4 = types.InlineKeyboardButton('ЧТ', callback_data='Четверг')
    btn5 = types.InlineKeyboardButton('ПТ', callback_data='Пятница')
    btn6 = types.InlineKeyboardButton('СБ', callback_data='Суббота')
    markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
    msg = bot.send_message(msg.chat.id, "Выбери день недели: \n", reply_markup=markup)
    # msg = bot.edit_message_text(text="Выбери день недели: \n", chat_id = msg.chat.id, message_id = msg.message_id, reply_markup=markup)
    return msg

def repbtns(chatid):                            # reply buttons
    repmarkup = types.ReplyKeyboardMarkup()
    repmarkup.row('Расписание полностью')
    repmarkup.row('Расписание четной недели')
    repmarkup.row('Расписание нечетной недели')
    repmarkup.row('Расписание по дням')
    repmarkup.row('Изменить параметры')
    repmarkup.row('Начать сначала')
    # repmarkup.row('Обновить расписание')

    bot.send_message(chatid, "Выбери формат вывода расписания\n", reply_markup=repmarkup)

def clear_chat(obj, msg):
    n = obj.msg_count
    for i in range(n-6):
        bot.delete_message(msg.chat.id, msg.id-i)
    obj.clear_attributes()
    button_message(msg)

def inline_btns_group(obj, chatid):
    obj.get_list_group()                        # формируем лист групп
    markup = types.InlineKeyboardMarkup()
    i = 0                                       # счетчик групп для колбека

    for group in obj.list_groups:
        if str(group) == 'None': continue

        btn = types.InlineKeyboardButton(text=group, callback_data=f'{i}group')
        i += 1
        markup.add(btn)
    bot.send_message(chatid, "Отлично! Теперь нужно выбрать группу", reply_markup=markup)

def inline_btns_edup(obj, chatid):
    obj.get_list_edup()                         # формируем лист обр программ
    markup = types.InlineKeyboardMarkup()
    i = 0                                       # счетчик обр программ для колбека

    for edup in obj.list_edup:
        edup = str(edup)
        if edup == 'None': continue
        btn = types.InlineKeyboardButton(text=edup, callback_data=f'{i}edup')
        i += 1
        markup.add(btn)
    bot.send_message(chatid, f"Образовательные программы {obj.napr.capitalize()}:", reply_markup=markup)

def inline_btns_napr(obj, chatid):
    obj.get_list_napr()                         # формируем лист направлений
    markup = types.InlineKeyboardMarkup()
    i = 0
    for name in obj.list_napr:
        if str(name) == 'None': continue
        btn = types.InlineKeyboardButton(text=name, callback_data=f'{i}napr')
        i += 1
        markup.add(btn)

    bot.send_message(chatid, f"Направления {obj.inst}:", reply_markup=markup)

def inline_btns_inst(obj, chatid, msg=''):

    obj.get_list_inst()                         # формируем лист направлений
    markup = types.InlineKeyboardMarkup()
    i = 0
    for name in obj.list_insts:
        if str(name) == 'None': continue
        btn = types.InlineKeyboardButton(text=name, callback_data=f'{i}inst')
        i += 1
        markup.add(btn)

    if msg != '':
        bot.edit_message_text(text=f"Институты {obj.course} курса:", chat_id = chatid, message_id = msg.message_id, reply_markup=markup)

    else:
        bot.send_message(chatid, f"Институты {obj.course} курса:", reply_markup=markup)

bot = telebot.TeleBot(config.token)
user_direct = Direct()                            #создаем объект нашего класса

@bot.message_handler(commands=['start'])
def button_message(message):
    # markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup = types.InlineKeyboardMarkup()
    btn = types.InlineKeyboardButton("Выбрать курс", callback_data = 'start')
    markup.add(btn)
    user_direct.msg_count = 1
    user_direct.add_token(message.chat.id)

    if user_direct.course == '': #проверяем на первый запуск
        bot.send_message(message.chat.id,'Привет! \nЯ бот-хранитель распиcания бакалавриата ГУУ!\n'
                                     'Сейчас нужно будет выбрать курс, затем появится список институтов.', reply_markup = markup)

    else:
        user_direct.clear_attributes()
        msg = bot.send_message(message.chat.id, 'Загрузка..', reply_markup = ReplyKeyboardRemove())
        bot.delete_message(message.chat.id, msg.id)
        bot.send_message(text='Начнем сначала.\nВыбирай курс, потом институт.', chat_id = message.chat.id, reply_markup=markup)

@bot.callback_query_handler(func=lambda call:True)
def callback_query(call):
    req = call.data.split('_')
    user_direct.msg_count += 1

    if req[0] == 'start':
        # bot.delete_message(call.message.chat.id, call.message.message_id)
        bot.edit_message_text('Привет! \nЯ бот-хранитель твоего распиcания!', chat_id = call.message.chat.id, message_id=call.message.message_id)
        markup = types.InlineKeyboardMarkup()

        for i in range(1, 5):
            btn = types.InlineKeyboardButton(text=f'{i}-курс', callback_data=f'{i}course')
            markup.add(btn)
        bot.delete_message(call.message.chat.id, call.message.message_id)
        bot.send_message(call.message.chat.id, "Бакалавриат. Курсы: ", reply_markup=markup)

    # получили курс
    if req[0][1:] == 'course':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        msg = bot.send_message(call.message.chat.id, "Загрузка..")
        course = int(str(req[0])[0])
        if course != user_direct.course:
            user_direct.clear_attributes()
            user_direct.set_course(course)
            user_direct.first_start()

        inline_btns_inst(user_direct, call.message.chat.id, msg)

    # получили название института
    if req[0][1:] == 'inst':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        user_direct.set_inst(user_direct.list_insts[int(req[0][:1])])      # добавили в наш объект выбранный институт

        inline_btns_napr(user_direct, call.message.chat.id)

    # получили название направления
    if req[0][1:] == 'napr':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        user_direct.set_napr(user_direct.list_napr[int(req[0][:1])])
        user_direct.get_list_edup()                           # формируем лист обр программ
        markup = types.InlineKeyboardMarkup()
        i = 0                                               # счетчик обр программ для колбека

        for edup in user_direct.list_edup:
            edup = str(edup)
            if edup == 'None': continue
            btn = types.InlineKeyboardButton(text = edup, callback_data = f'{i}edup')
            i+=1
            markup.add(btn)
        bot.send_message(call.message.chat.id, f"Образовательные программы {user_direct.napr.capitalize()}:", reply_markup = markup)

    # получили название обр программы
    if req[0][1:] == 'edup':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        user_direct.set_edup(user_direct.list_edup[int(req[0][:1])])

        inline_btns_group(user_direct, call.message.chat.id)

    # получили группу
    if req[0][1:] == 'group':
        bot.delete_message(call.message.chat.id, call.message.message_id)
        user_direct.set_group(user_direct.list_groups[int(req[0][:1])])

        bot.send_message(call.message.chat.id, "Настройка завершена!✅\n"
                                               f"Курс: {user_direct.course} \n"
                                               f"Направление подготовки: {user_direct.napr} \n"
                                               f"Образовательная программа: {user_direct.edup} \n"
                                               f"Группа: {user_direct.group} \n")

        repbtns(call.message.chat.id)

    # ВЫВОД РАСПИСАНИЯ ==========================================================

    weekdays = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']

    if req[0] in weekdays:
        answer = user_direct.get_scd_weekday(req[0])
        # bot.send_message(call.message.chat.id, answer)
        user_direct.week_msg = bot.edit_message_text(text=answer, chat_id=call.message.chat.id, message_id=call.message.message_id)
        weekdayscd(call.message, user_direct)


@bot.message_handler(content_types=['text'])
def message_reply(message):
    user_direct.msg_count += 1
    chatid = message.chat.id

    if message.text == "Расписание полностью":
        fullsqd(user_direct, chatid)

    if message.text == "Расписание четной недели":
        evenscd(user_direct, 'Чёт.', chatid)

    if message.text == "Расписание нечетной недели":
        evenscd(user_direct, 'Нечёт.', chatid)

    if message.text == "Расписание по дням":
        weekdayscd(message, user_direct)

    if message.text == "Изменить параметры":
        markup = types.ReplyKeyboardMarkup(one_time_keyboard=True)
        markup.add('Изменить группу')
        markup.add('Изменить образовательную программу')
        markup.add('Изменить направление')
        markup.add('Изменить институт')
        bot.send_message(chatid, 'Выбери, что хочешь изменить', reply_markup=markup)

    if message.text == "Начать сначала":
        # markup = types.ReplyKeyboardMarkup(one_time_keyboard=True)
        # markup.add('Сотри все')
        # markup.add('Оставь')
        # bot.send_message(chatid, 'Могу стереть все предыдущие сообщения?', reply_markup = markup)

        button_message(message)

    # if message.text.lower() == 'Сотри все'.lower():
    #     clear_chat(my_direct, message)
    #
    # if message.text.lower() == 'Оставь'.lower():
    #     button_message(message) # запуск с команды старт

    if message.text.lower() == 'Обновить расписание'.lower():
        msg = bot.send_message(message.chat.id, "Загрузка..")
        # os.remove(my_direct.path)
        user_direct.first_start()         # тк курс у нас не меняется, то мы просто подгружаем и обрабатываем док с сайта
        bot.edit_message_text(f'Распсиание для {user_direct.course} курса обновлено.', chat_id = message.chat.id, message_id = msg.message_id)
        repbtns(message.chat.id)
        pass

    if message.text.lower() == 'Изменить группу'.lower():
        user_direct.group = ''
        inline_btns_group(user_direct, message.chat.id)

    if message.text.lower() == 'Изменить образовательную программу'.lower():
        user_direct.group = ''
        user_direct.list_groups = ''
        user_direct.edup = ''
        inline_btns_edup(user_direct, message.chat.id)

    if message.text.lower() == 'Изменить направление'.lower():
        user_direct.group = ''
        user_direct.list_groups = ''
        user_direct.edup = ''
        user_direct.list_edup = ''
        user_direct.napr = ''
        inline_btns_napr(user_direct, message.chat.id)


    if message.text.lower() == 'Изменить институт'.lower():
        user_direct.group = ''
        user_direct.list_groups = ''
        user_direct.edup = ''
        user_direct.list_edup = ''
        user_direct.napr = ''
        user_direct.list_napr = ''
        user_direct.inst = ''
        inline_btns_inst(user_direct, message.chat.id)


bot.infinity_polling()
