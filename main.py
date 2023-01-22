import telebot
from telebot import types
import datetime
import xlrd
from datetime import timedelta
import requests
from bs4 import BeautifulSoup
import time

bot = telebot.TeleBot("5951494078:AAF5YF_v358mhS9OY6iQqgdX1IP-FuAyW5k")
b = []
test_chanel = "@RaspisniTest"
url = "https://school-200.ru"
chatId = 751762217
d = 0




def check_date():
    hours = datetime.datetime.now() + datetime.timedelta(hours=5)
    now = datetime.datetime.now() + datetime.timedelta(hours=5)
    d = datetime.datetime.today() + datetime.timedelta(hours=5)
    if now.weekday()+1 <= 5:
        if hours.hour > 16:
            day = d + datetime.timedelta(days=1)
        else:
            day = d
    if now.weekday()+1 == 6:
        day = d + datetime.timedelta(days=2)
    if now.weekday()+1 == 7:
        day = d + datetime.timedelta(days=1)




    dls = f"https://school-200.ru/doc/schedule/2022-2023/changes_{day.day}-{day.month}-{day.year}_5-11.xls"
    dls1 = f"https://school-200.ru/doc/schedule/2022-2023/changes_{day.day}-{day.month}-{day.year}_1-4.xls"
    resp = requests.get(dls)
    resp1 = requests.get(dls1)

    try:
        output = open('test.xls', 'wb')
        output.write(resp.content)
        output.close()

        output1 = open('test1.xls', 'wb')
        output1.write(resp1.content)
        output1.close()

        path = "test.xls"
        path1 = "test1.xls"

        inputWorkbook = xlrd.open_workbook(path, formatting_info=True)
        inputWorksheet = inputWorkbook.sheet_by_index(0)
        sheets = inputWorkbook.sheet_names()

        inputWorkbook1 = xlrd.open_workbook(path1, formatting_info=True)
        inputWorksheet1 = inputWorkbook1.sheet_by_index(0)
        sheets1 = inputWorkbook1.sheet_names()

        return [inputWorkbook, inputWorksheet, sheets, inputWorkbook1, inputWorksheet1, sheets1, day]

    except xlrd.biffh.XLRDError:
        return True
@bot.message_handler(commands=['start'])
def start(message):
    now = datetime.datetime.now()
    markup = types.InlineKeyboardMarkup(row_width=2)
    btn_1 = types.InlineKeyboardButton(text=f'Первая смена', callback_data='btn1')
    btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
    markup.add(btn_1, btn_2)

    bot.send_message(message.chat.id,
                     f"Привет, я бот который буду присылать тебе расписание. 👋 \nДля начала скажи в какую смену ты учишься? 🤔",
                     reply_markup=markup)
    d = 0























@bot.callback_query_handler(func=lambda call: True)
def check_callback_data(call):
    if call.data == "btn1":
        hours = datetime.datetime.now() + datetime.timedelta(hours=5)
        hours1 = datetime.datetime.today() + datetime.timedelta(hours=5)
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='Проверить новое расписание', callback_data='check_raspin')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"Расписание на следующий день еще не вышло :(", reply_markup=markup)
        else:
            hours1 = datetime.datetime.today()
            markup = types.InlineKeyboardMarkup(row_width=5)
            btn_bg21 = types.InlineKeyboardButton(text='1а', callback_data="btn11")
            btn_bg22 = types.InlineKeyboardButton(text='1б', callback_data="btn12")
            btn_bg23 = types.InlineKeyboardButton(text='1в', callback_data="btn13")
            btn_bg24 = types.InlineKeyboardButton(text='1г', callback_data="btn14")
            btn_bg25 = types.InlineKeyboardButton(text='1д', callback_data="btn15")
            btn_bg26 = types.InlineKeyboardButton(text='2б', callback_data="btn21")
            btn_bg27 = types.InlineKeyboardButton(text='4а', callback_data="btn41")
            btn_bg28 = types.InlineKeyboardButton(text='4б', callback_data="btn42")
            btn_bg29 = types.InlineKeyboardButton(text='4г', callback_data="btn43")
            btn_bg30 = types.InlineKeyboardButton(text='4е', callback_data="btn44")
            btn_bg1 = types.InlineKeyboardButton(text='5а', callback_data="btn51")
            btn_bg2 = types.InlineKeyboardButton(text='5б', callback_data="btn52")
            btn_bg3 = types.InlineKeyboardButton(text='5в', callback_data="btn53")
            btn_bg4 = types.InlineKeyboardButton(text='5г', callback_data="btn54")
            btn_bg5 = types.InlineKeyboardButton(text='5д', callback_data="btn55")
            btn_bg6 = types.InlineKeyboardButton(text='7а', callback_data="btn71")
            btn_bg7 = types.InlineKeyboardButton(text='7б', callback_data="btn72")
            btn_bg8 = types.InlineKeyboardButton(text='7в', callback_data="btn73")
            btn_bg9 = types.InlineKeyboardButton(text='7г', callback_data="btn74")
            btn_bg10 = types.InlineKeyboardButton(text='8в', callback_data="btn8")
            btn_bg11 = types.InlineKeyboardButton(text='9а', callback_data="btn91")
            btn_bg12 = types.InlineKeyboardButton(text='9б', callback_data="btn92")
            btn_bg13 = types.InlineKeyboardButton(text='9в', callback_data="btn93")
            btn_bg14 = types.InlineKeyboardButton(text='9г', callback_data="btn94")
            btn_bg15 = types.InlineKeyboardButton(text='10а', callback_data="btn101")
            btn_bg16 = types.InlineKeyboardButton(text='10б', callback_data="btn102")
            btn_bg17 = types.InlineKeyboardButton(text='10в', callback_data="btn103")
            btn_bg18 = types.InlineKeyboardButton(text='10г', callback_data="btn104")
            btn_bg19 = types.InlineKeyboardButton(text='11а', callback_data="btn111")
            btn_bg20 = types.InlineKeyboardButton(text='11б', callback_data="btn112")
            markup.add(btn_bg21,
                       btn_bg22,
                       btn_bg23,
                       btn_bg24,
                       btn_bg25,
                       btn_bg26,
                       btn_bg27,
                       btn_bg28,
                       btn_bg29,
                       btn_bg30, btn_bg1, btn_bg2, btn_bg3, btn_bg4, btn_bg5, btn_bg6, btn_bg7, btn_bg8,
                       btn_bg9,
                       btn_bg10,
                       btn_bg11,
                       btn_bg12,
                       btn_bg13,
                       btn_bg14,
                       btn_bg15,
                       btn_bg16,
                       btn_bg17,
                       btn_bg18,
                       btn_bg19,
                       btn_bg20)
            markup.add(telebot.types.InlineKeyboardButton('Вторая смена', callback_data='btn2'))
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"Выбери класс где ты учишься.", reply_markup=markup)

    if call.data == "btn1":
        hours = datetime.datetime.now() + datetime.timedelta(hours=5)
        hours1 = datetime.datetime.today() + datetime.timedelta(hours=5)
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='Проверить на наличие нового расписания', callback_data='check_raspin')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"Расписание на следующий день еще не вышло :(", reply_markup=markup)
        else:
            hours1 = datetime.datetime.today()
            markup = types.InlineKeyboardMarkup(row_width=5)
            btn_bg21 = types.InlineKeyboardButton(text='2а', callback_data="btn22")
            btn_bg22 = types.InlineKeyboardButton(text='2в', callback_data="btn23")
            btn_bg23 = types.InlineKeyboardButton(text='2г', callback_data="btn24")
            btn_bg24 = types.InlineKeyboardButton(text='2д', callback_data="btn25")
            btn_bg25 = types.InlineKeyboardButton(text='3а', callback_data="btn31")
            btn_bg26 = types.InlineKeyboardButton(text='3б', callback_data="btn32")
            btn_bg27 = types.InlineKeyboardButton(text='3в', callback_data="btn33")
            btn_bg28 = types.InlineKeyboardButton(text='3г', callback_data="btn34")
            btn_bg29 = types.InlineKeyboardButton(text='3д', callback_data="btn35")
            btn_bg30 = types.InlineKeyboardButton(text='4в', callback_data="btn43")
            btn_bg1 = types.InlineKeyboardButton(text='4д', callback_data="btn45")
            btn_bg2 = types.InlineKeyboardButton(text='6а', callback_data="btn61")
            btn_bg3 = types.InlineKeyboardButton(text='6б', callback_data="btn62")
            btn_bg4 = types.InlineKeyboardButton(text='6в', callback_data="btn63")
            btn_bg5 = types.InlineKeyboardButton(text='6г', callback_data="btn64")
            btn_bg6 = types.InlineKeyboardButton(text='6д', callback_data="btn65")
            btn_bg7 = types.InlineKeyboardButton(text='8а', callback_data="btn81")
            btn_bg8 = types.InlineKeyboardButton(text='8б', callback_data="btn82")
            btn_bg9 = types.InlineKeyboardButton(text='8г', callback_data="btn84")

            markup.add(btn_bg21,
                       btn_bg22,
                       btn_bg23,
                       btn_bg24,
                       btn_bg25,
                       btn_bg26,
                       btn_bg27,
                       btn_bg28,
                       btn_bg29,
                       btn_bg30, btn_bg1, btn_bg2, btn_bg3, btn_bg4, btn_bg5, btn_bg6, btn_bg7, btn_bg8,
                       btn_bg9)
            markup.add(telebot.types.InlineKeyboardButton('Первая смена', callback_data='btn1'))
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text="Выбери класс где ты учишься.", reply_markup=markup)

    elif call.data == "check_raspin1":
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='Проверить на наличие нового расписания', callback_data='check_raspin')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"Расписание на следующий день еще не вышло :(", reply_markup=markup)
        else:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text=f'Первая смена', callback_data='btn1')
            btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
            markup.add(btn_1, btn_2)

            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"Cкажи в какую смену ты учишься? 🤔",
                             reply_markup=markup)

    elif call.data == "check_raspin":
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='Проверить новое расписание', callback_data='check_raspin1')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"Расписание на следующий день еще не вышло :(", reply_markup=markup)
        else:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text=f'Первая смена', callback_data='btn1')
            btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
            markup.add(btn_1, btn_2)

            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"Cкажи в какую смену ты учишься? 🤔",
                                  reply_markup=markup)
    elif call.data == "btn22":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(28, 2)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>2а-ДО</b>')
        else:
            b.append('<b>2а</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(29, 34):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - Дистант")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)
        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')










    elif call.data == "btn23":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(28, 3)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>2в-ДО</b>')
        else:
            b.append('<b>2в</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(29, 34):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - Дистант")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn24":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(28, 4)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>2в-ДО</b>')
        else:
            b.append('<b>2г</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(29, 34):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - Дистант")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn25":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(28, 5)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>2д-ДО</b>')
        else:
            b.append('<b>2д</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(29, 34):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - Дистант")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn31":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(35, 2)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>3а-ДО</b>')
        else:
            b.append('<b>3а</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(36, 41):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - Дистант")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn33":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(35, 4)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>3в-ДО</b>')
        else:
            b.append('<b>3в</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(36, 41):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - Дистант")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn34":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []

        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(35, 5)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>3г-ДО</b>')
        else:
            b.append('<b>3г</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(36, 41):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - Дистант")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn32":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(35, 3)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>3б-ДО</b>')
        else:
            b.append('<b>3б</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(36, 41):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - Дистант")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet1.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn35":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(35, 6)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>3д-ДО</b>')
        else:
            b.append('<b>3д</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(36, 41):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)} - Дистант")
            elif inputWorksheet1.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn43":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(18, 5)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>4г-ДО</b>')
        else:
            b.append('<b>4г</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(18, 24):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - Дистант")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn45":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 6
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(42, 3)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>4д-ДО</b>')
        else:
            b.append('<b>4д</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 50, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 14, 30, 00)
        for i in range(43, 48):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - Дистант")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn61":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 5
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(53, 2)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>6а-ДО</b>')
        else:
            b.append('<b>6а</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 13, 40, 00)
        for i in range(54, 61):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - Дистант")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn62":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 5
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(40, 4)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>6б-ДО</b>')
        else:
            b.append('<b>6б</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 13, 40, 00)
        for i in range(54, 61):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - Дистант")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn63":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 5
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(53, 4)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>6в-ДО</b>')
        else:
            b.append('<b>6в</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 13, 40, 00)
        for i in range(54, 61):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - Дистант")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn64":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 5
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(53, 5)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>6г-ДО</b>')
        else:
            b.append('<b>6г</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 13, 40, 00)
        for i in range(54, 61):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - Дистант")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn65":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 5
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(53, 6)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>6д-ДО</b>')
        else:
            b.append('<b>6д</b>')
        lessons = datetime.datetime(2022, 12, 3, 13, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 13, 40, 00)
        for i in range(54, 61):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - Дистант")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn81":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 4
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(62, 2)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>8а-ДО</b>')
        else:
            b.append('<b>8а</b>')
        lessons = datetime.datetime(2022, 12, 3, 12, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 12, 40, 00)
        for i in range(63, 71):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - Дистант")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn82":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 4
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(62, 3)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>8б-ДО</b>')
        else:
            b.append('<b>8б</b>')
        lessons = datetime.datetime(2022, 12, 3, 12, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 12, 40, 00)
        for i in range(63, 71):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - Дистант")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn84":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 4
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(62, 4)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>8г-ДО</b>')
        else:
            b.append('<b>8г</b>')
        lessons = datetime.datetime(2022, 12, 3, 12, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 12, 40, 00)
        for i in range(63, 71):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - Дистант")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn11":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 2)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>1а-ДО</b>')
        else:
            b.append('<b>1а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 16):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - Дистант")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)
        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn12":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 3)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>1б-ДО</b>')
        else:
            b.append('<b>1б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 16):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - Дистант")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)
        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn13":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 4)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>1в-ДО</b>')
        else:
            b.append('<b>1в</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 16):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - Дистант")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn14":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 5)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>1г-ДО</b>')
        else:
            b.append('<b>1г</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 16):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - Дистант")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn15":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 6)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>1д-ДО</b>')
        else:
            b.append('<b>1д</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 16):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)} - Дистант")
            elif inputWorksheet1.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn21":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(17, 2)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>2б-ДО</b>')
        else:
            b.append('<b>2б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(18, 24):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - Дистант")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn41":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(17, 3)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>4а-ДО</b>')
        else:
            b.append('<b>4а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(18, 24):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - Дистант")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')




    elif call.data == "btn42":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(17, 4)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>4б-ДО</b>')
        else:
            b.append('<b>4б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(18, 24):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - Дистант")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn43":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(17, 5)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>4г-ДО</b>')
        else:
            b.append('<b>4г</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(18, 24):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - Дистант")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn44":
        check = check_date()
        sheets = check[2]
        inputWorkbook1 = check[3]
        inputWorksheet1 = check[4]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook1.sheet_by_index(index)
            xfx = sheet.cell_xf_index(17, 6)
            xf = inputWorkbook1.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>4е-ДО</b>')
        else:
            b.append('<b>4е</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(18, 24):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook1.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook1.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)} - Дистант")
            elif inputWorksheet1.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn51":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 2)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>5а-ДО</b>')
        else:
            b.append('<b>5а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 18):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - Дистант")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn52":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 3)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>5б-ДО</b>')
        else:
            b.append('<b>5б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 18):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - Дистант")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn53":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 4)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>5в-ДО</b>')
        else:
            b.append('<b>5в</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 18):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - Дистант")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(telebot.types.InlineKeyboardButton(text='Первая смена', callback_data='btn1'))
        markup.add(telebot.types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2'))
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                text=b, reply_markup=markup)

    elif call.data == "btn54":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 5)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>5г-ДО</b>')
        else:
            b.append('<b>5г</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 18):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - Дистант")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(telebot.types.InlineKeyboardButton(text='Первая смена', callback_data='btn1'))
        markup.add(telebot.types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2'))
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                text=b, reply_markup=markup)

    elif call.data == "btn55":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(10, 6)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>5д-ДО</b>')
        else:
            b.append('<b>5д</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(11, 18):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - Дистант")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(telebot.types.InlineKeyboardButton(text='Первая смена', callback_data='btn1'))
        markup.add(telebot.types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2'))
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                text=b, reply_markup=markup)

    elif call.data == "btn71":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(19, 2)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>7а-ДО</b>')
        else:
            b.append('<b>7а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(20, 29):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - Дистант")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn72":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(19, 3)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>7б-ДО</b>')
        else:
            b.append('<b>7б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(20, 29):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - Дистант")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn73":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(19, 4)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>7в-ДО</b>')
        else:
            b.append('<b>7в</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(20, 29):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - Дистант")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn74":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(19, 5)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>2в-ДО</b>')
        else:
            b.append('7г')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(20, 29):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - Дистант")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn8":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(19, 6)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>8а-ДО</b>')
        else:
            b.append('<b>8а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(20, 29):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - Дистант")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn91":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(30, 2)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>9а-ДО</b>')
        else:
            b.append('<b>9а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(31, 39):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - Дистант")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn92":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(30, 3)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>9б-ДО</b>')
        else:
            b.append('<b>9б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(31, 39):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - Дистант")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn93":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(30, 4)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>9в-ДО</b>')
        else:
            b.append('<b>9в</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(31, 39):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - Дистант")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn94":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(30, 5)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>9г-ДО</b>')
        else:
            b.append('<b>9г</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(31, 39):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - Дистант")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn101":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(30, 6)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>10а-ДО</b>')
        else:
            b.append('<b>10а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(31, 39):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - Дистант")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn102":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(40, 2)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>10б-ДО</b>')
        else:
            b.append('<b>10б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(41, 49):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 2)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - Дистант")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 2)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn103":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(40, 3)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>10в-ДО</b>')
        else:
            b.append('<b>10в</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(41, 49):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 3)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - Дистант")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 3)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn104":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(40, 4)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>10г-ДО</b>')
        else:
            b.append('<b>10г</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(41, 49):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 4)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - Дистант")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 4)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')



    elif call.data == "btn111":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(40, 5)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>11а-ДО</b>')
        else:
            b.append('<b>11а</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(41, 49):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 5)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - Дистант")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 5)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


    elif call.data == "btn112":
        check = check_date()
        sheets = check[2]
        inputWorkbook = check[0]
        inputWorksheet = check[1]
        a = 0
        b = []
        for index, sh in enumerate(sheets):
            sheet = inputWorkbook.sheet_by_index(index)
            xfx = sheet.cell_xf_index(40, 6)
            xf = inputWorkbook.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
        if bgx == 13:
            b.append('<b>11б-ДО</b>')
        else:
            b.append('<b>11б</b>')
        lessons = datetime.datetime(2022, 12, 3, 8, 00, 00)
        lessons1 = datetime.datetime(2022, 12, 3, 8, 40, 00)
        for i in range(41, 49):
            a += 1
            b.append('')
            for index, sh in enumerate(sheets):
                sheet = inputWorkbook.sheet_by_index(index)
                xfx = sheet.cell_xf_index(i, 6)
                xf = inputWorkbook.xf_list[xfx]
                bgx = xf.background.pattern_colour_index

            if bgx == 36:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - Дистант")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} Урока нет 🚫")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 6)}   ⏰ {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='Первая смена', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='Вторая смена', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


bot.polling(non_stop=True)