import config as cf
import telebot
from telebot import types
from telebot.types import InlineKeyboardButton

from main import *

# token = '5679216888:AAEnHl7wKQmR4mXwrqWQQIVGVztbqtINeBQ'
bot = telebot.TeleBot(cf.token)

# @bot.message_handler(commands=['start'])
# def start_message(message):
#     bot.send_message(message.chat.id,'Привет')

@bot.message_handler(commands=['start'])
def button_message(message):
    markup=types.ReplyKeyboardMarkup(resize_keyboard=True)
    item2 = types.KeyboardButton("Выбрать институт")
    markup.add(item2)
    bot.send_message(message.chat.id,'Привет! \nЯ бот-хранитель твоего распиания! \nДля начала, нужно выбрать институт',reply_markup=markup)

@bot.message_handler(content_types='text')
def message_reply(message):
    if message.text=="Кнопка":
        bot.send_message(message.chat.id,"https://habr.com/ru/users/lubaznatel/")
    if message.text=="Выбрать институт":
        markup = types.InlineKeyboardMarkup()
        buttons = []
        for name in full_inst_list:
            # buttons.append(types.InlineKeyboardButton(text = str(name), callback_data='callback'))
            # markup.row(*buttons)
            btn = types.InlineKeyboardButton(text = str(name), callback_data='callback')
            markup.add(btn)
        bot.send_message(message.chat.id, "Вот что есть", reply_markup=markup)

bot.infinity_polling()
