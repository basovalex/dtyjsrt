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
    den = day.day
    if int(den) < 10:
        den = '0' + str(day.day)
    month = day.month
    if day.month < 10:
        month = '0' + str(day.month)





    dls = f"https://school-200.ru/doc/schedule/2022-2023/changes_{den}-{month}-{day.year}_5-11.xls"
    dls1 = f"https://school-200.ru/doc/schedule/2022-2023/changes_{den}-{month}-{day.year}_1-4.xls"
    resp = requests.get(dls)
    resp1 = requests.get(dls1)
    print(dls)

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
    btn_1 = types.InlineKeyboardButton(text=f'???????????? ??????????', callback_data='btn1')
    btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
    markup.add(btn_1, btn_2)

    bot.send_message(message.chat.id,
                     f"????????????, ?? ?????? ?????????????? ???????? ?????????????????? ???????? ????????????????????. ???? \n?????? ???????????? ?????????? ?? ?????????? ?????????? ???? ??????????????? ????",
                     reply_markup=markup)
    d = 0

@bot.callback_query_handler(func=lambda call: True)
def check_callback_data(call):
    if call.data == "btn1":
        hours = datetime.datetime.now() + datetime.timedelta(hours=5)
        hours1 = datetime.datetime.today() + datetime.timedelta(hours=5)
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='?????????????????? ?????????? ????????????????????', callback_data='check_raspin')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"???????????????????? ???? ?????????????????? ???????? ?????? ???? ?????????? :(", reply_markup=markup)
        else:
            hours1 = datetime.datetime.today()
            markup = types.InlineKeyboardMarkup(row_width=5)
            btn_bg21 = types.InlineKeyboardButton(text='1??', callback_data="btn11")
            btn_bg22 = types.InlineKeyboardButton(text='1??', callback_data="btn12")
            btn_bg23 = types.InlineKeyboardButton(text='1??', callback_data="btn13")
            btn_bg24 = types.InlineKeyboardButton(text='1??', callback_data="btn14")
            btn_bg25 = types.InlineKeyboardButton(text='1??', callback_data="btn15")
            btn_bg26 = types.InlineKeyboardButton(text='2??', callback_data="btn21")
            btn_bg27 = types.InlineKeyboardButton(text='4??', callback_data="btn41")
            btn_bg28 = types.InlineKeyboardButton(text='4??', callback_data="btn42")
            btn_bg29 = types.InlineKeyboardButton(text='4??', callback_data="btn43")
            btn_bg30 = types.InlineKeyboardButton(text='4??', callback_data="btn44")
            btn_bg1 = types.InlineKeyboardButton(text='5??', callback_data="btn51")
            btn_bg2 = types.InlineKeyboardButton(text='5??', callback_data="btn52")
            btn_bg3 = types.InlineKeyboardButton(text='5??', callback_data="btn53")
            btn_bg4 = types.InlineKeyboardButton(text='5??', callback_data="btn54")
            btn_bg5 = types.InlineKeyboardButton(text='5??', callback_data="btn55")
            btn_bg6 = types.InlineKeyboardButton(text='7??', callback_data="btn71")
            btn_bg7 = types.InlineKeyboardButton(text='7??', callback_data="btn72")
            btn_bg8 = types.InlineKeyboardButton(text='7??', callback_data="btn73")
            btn_bg9 = types.InlineKeyboardButton(text='7??', callback_data="btn74")
            btn_bg10 = types.InlineKeyboardButton(text='8??', callback_data="btn8")
            btn_bg11 = types.InlineKeyboardButton(text='9??', callback_data="btn91")
            btn_bg12 = types.InlineKeyboardButton(text='9??', callback_data="btn92")
            btn_bg13 = types.InlineKeyboardButton(text='9??', callback_data="btn93")
            btn_bg14 = types.InlineKeyboardButton(text='9??', callback_data="btn94")
            btn_bg15 = types.InlineKeyboardButton(text='10??', callback_data="btn101")
            btn_bg16 = types.InlineKeyboardButton(text='10??', callback_data="btn102")
            btn_bg17 = types.InlineKeyboardButton(text='10??', callback_data="btn103")
            btn_bg18 = types.InlineKeyboardButton(text='10??', callback_data="btn104")
            btn_bg19 = types.InlineKeyboardButton(text='11??', callback_data="btn111")
            btn_bg20 = types.InlineKeyboardButton(text='11??', callback_data="btn112")
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
            markup.add(telebot.types.InlineKeyboardButton('???????????? ??????????', callback_data='btn2'))
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"???????????? ?????????? ?????? ???? ??????????????.", reply_markup=markup)

    if call.data == "btn2":
        hours = datetime.datetime.now() + datetime.timedelta(hours=5)
        hours1 = datetime.datetime.today() + datetime.timedelta(hours=5)
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='?????????????????? ???? ?????????????? ???????????? ????????????????????', callback_data='check_raspin')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"???????????????????? ???? ?????????????????? ???????? ?????? ???? ?????????? :(", reply_markup=markup)
        else:
            hours1 = datetime.datetime.today()
            markup = types.InlineKeyboardMarkup(row_width=5)
            btn_bg21 = types.InlineKeyboardButton(text='2??', callback_data="btn22")
            btn_bg22 = types.InlineKeyboardButton(text='2??', callback_data="btn23")
            btn_bg23 = types.InlineKeyboardButton(text='2??', callback_data="btn24")
            btn_bg24 = types.InlineKeyboardButton(text='2??', callback_data="btn25")
            btn_bg25 = types.InlineKeyboardButton(text='3??', callback_data="btn31")
            btn_bg26 = types.InlineKeyboardButton(text='3??', callback_data="btn32")
            btn_bg27 = types.InlineKeyboardButton(text='3??', callback_data="btn33")
            btn_bg28 = types.InlineKeyboardButton(text='3??', callback_data="btn34")
            btn_bg29 = types.InlineKeyboardButton(text='3??', callback_data="btn35")
            btn_bg30 = types.InlineKeyboardButton(text='4??', callback_data="btn43")
            btn_bg1 = types.InlineKeyboardButton(text='4??', callback_data="btn45")
            btn_bg2 = types.InlineKeyboardButton(text='6??', callback_data="btn61")
            btn_bg3 = types.InlineKeyboardButton(text='6??', callback_data="btn62")
            btn_bg4 = types.InlineKeyboardButton(text='6??', callback_data="btn63")
            btn_bg5 = types.InlineKeyboardButton(text='6??', callback_data="btn64")
            btn_bg6 = types.InlineKeyboardButton(text='6??', callback_data="btn65")
            btn_bg7 = types.InlineKeyboardButton(text='8??', callback_data="btn81")
            btn_bg8 = types.InlineKeyboardButton(text='8??', callback_data="btn82")
            btn_bg9 = types.InlineKeyboardButton(text='8??', callback_data="btn84")

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
            markup.add(telebot.types.InlineKeyboardButton('???????????? ??????????', callback_data='btn1'))
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text="???????????? ?????????? ?????? ???? ??????????????.", reply_markup=markup)

    elif call.data == "check_raspin1":
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='?????????????????? ???? ?????????????? ???????????? ????????????????????', callback_data='check_raspin')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"???????????????????? ???? ?????????????????? ???????? ?????? ???? ?????????? :(", reply_markup=markup)
        else:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text=f'???????????? ??????????', callback_data='btn1')
            btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
            markup.add(btn_1, btn_2)

            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"C???????? ?? ?????????? ?????????? ???? ??????????????? ????",
                             reply_markup=markup)

    elif call.data == "check_raspin":
        if check_date() == True:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text='?????????????????? ?????????? ????????????????????', callback_data='check_raspin1')
            markup.add(btn_1)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"???????????????????? ???? ?????????????????? ???????? ?????? ???? ?????????? :(", reply_markup=markup)
        else:
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn_1 = types.InlineKeyboardButton(text=f'???????????? ??????????', callback_data='btn1')
            btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
            markup.add(btn_1, btn_2)

            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                                  text=f"C???????? ?? ?????????? ?????????? ???? ??????????????? ????",
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
            b.append('<b>2??-????</b>')
        else:
            b.append('<b>2??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)
        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>2??-????</b>')
        else:
            b.append('<b>2??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>2??-????</b>')
        else:
            b.append('<b>2??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>2??-????</b>')
        else:
            b.append('<b>2??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>3??-????</b>')
        else:
            b.append('<b>3??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>3??-????</b>')
        else:
            b.append('<b>3??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>3??-????</b>')
        else:
            b.append('<b>3??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>3??-????</b>')
        else:
            b.append('<b>3??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet1.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>3??-????</b>')
        else:
            b.append('<b>3??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>4??-????</b>')
        else:
            b.append('<b>4??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>4??-????</b>')
        else:
            b.append('<b>4??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6 or a == 9 or a == 10:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>6??-????</b>')
        else:
            b.append('<b>6??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>6??-????</b>')
        else:
            b.append('<b>6??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>6??-????</b>')
        else:
            b.append('<b>6??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>6??-????</b>')
        else:
            b.append('<b>6??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>6??-????</b>')
        else:
            b.append('<b>6??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>8??-????</b>')
        else:
            b.append('<b>8??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>8??-????</b>')
        else:
            b.append('<b>8??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>8??-????</b>')
        else:
            b.append('<b>8??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6  or a == 9 or a == 10 or a == 11:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>1??-????</b>')
        else:
            b.append('<b>1??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)
        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>1??-????</b>')
        else:
            b.append('<b>1??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)
        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>1??-????</b>')
        else:
            b.append('<b>1??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>1??-????</b>')
        else:
            b.append('<b>1??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>1??-????</b>')
        else:
            b.append('<b>1??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>2??-????</b>')
        else:
            b.append('<b>2??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>4??-????</b>')
        else:
            b.append('<b>4??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>4??-????</b>')
        else:
            b.append('<b>4??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>4??-????</b>')
        else:
            b.append('<b>4??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>4??-????</b>')
        else:
            b.append('<b>4??</b>')
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
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet1.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet1.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>5??-????</b>')
        else:
            b.append('<b>5??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>5??-????</b>')
        else:
            b.append('<b>5??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>5??-????</b>')
        else:
            b.append('<b>5??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(telebot.types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1'))
        markup.add(telebot.types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2'))
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
            b.append('<b>5??-????</b>')
        else:
            b.append('<b>5??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(telebot.types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1'))
        markup.add(telebot.types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2'))
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
            b.append('<b>5??-????</b>')
        else:
            b.append('<b>5??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(telebot.types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1'))
        markup.add(telebot.types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2'))
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
            b.append('<b>7??-????</b>')
        else:
            b.append('<b>7??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>7??-????</b>')
        else:
            b.append('<b>7??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>7??-????</b>')
        else:
            b.append('<b>7??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>2??-????</b>')
        else:
            b.append('7??')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>8??-????</b>')
        else:
            b.append('<b>8??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>9??-????</b>')
        else:
            b.append('<b>9??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>9??-????</b>')
        else:
            b.append('<b>9??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>9??-????</b>')
        else:
            b.append('<b>9??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>9??-????</b>')
        else:
            b.append('<b>9??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>10??-????</b>')
        else:
            b.append('<b>10??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>10??-????</b>')
        else:
            b.append('<b>10??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 2)} - ??????????????")
            elif inputWorksheet.cell_value(i, 2) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 2)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>10??-????</b>')
        else:
            b.append('<b>10??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 3)} - ??????????????")
            elif inputWorksheet.cell_value(i, 3) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 3)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>10??-????</b>')
        else:
            b.append('<b>10??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 4)} - ??????????????")
            elif inputWorksheet.cell_value(i, 4) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 4)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>11??-????</b>')
        else:
            b.append('<b>11??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 5)} - ??????????????")
            elif inputWorksheet.cell_value(i, 5) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 5)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
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
            b.append('<b>11??-????</b>')
        else:
            b.append('<b>11??</b>')
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
                b.append(f"{a} {inputWorksheet.cell_value(i, 6)} - ??????????????")
            elif inputWorksheet.cell_value(i, 6) == '':
                b.append(f"{a} ?????????? ?????? ????")
            else:
                b.append(
                    f"{a} {inputWorksheet.cell_value(i, 6)}   ??? {lessons.strftime('%H:%M')} - {lessons1.strftime('%H:%M')}")
            if a == 6:
                lessons = lessons1 + timedelta(minutes=10)
                lessons1 = lessons + timedelta(minutes=40)
            else:
                lessons = lessons1 + timedelta(minutes=20)
                lessons1 = lessons + timedelta(minutes=40)
        b = '\n'.join(b)

        markup = types.InlineKeyboardMarkup(row_width=2)
        btn_1 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn1')
        btn_2 = types.InlineKeyboardButton(text='???????????? ??????????', callback_data='btn2')
        markup.add(btn_1, btn_2)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.id,
                              text=b, reply_markup=markup, parse_mode='html')


bot.polling(non_stop=True)