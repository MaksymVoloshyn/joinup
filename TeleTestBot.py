import telebot;
import xlrd;

while True:

    bot = telebot.TeleBot('869444022:AAGC1z9ED1gmaV_lPZZRMyptOJody1l-xR0')
    pass_database = xlrd.open_workbook('C:/DBS.xls',formatting_info=True,on_demand = True)
    sheet = pass_database.sheet_by_index(0)
    last_row = sheet.nrows

    @bot.message_handler(content_types=['text'])
    def get_text_messages(message):
        correct = 0;
        if last_row > 0:
            for row in range(0, last_row):
                if (message.text == sheet.cell_value(row, 0)):
                    bot.send_message(-1001132150751, sheet.cell_value(row, 1))
                    correct = 1;
                    break
        if correct == 0:
            bot.send_message(-1001132150751, "Пароль не найден!")

    pass_database.release_resources()
    del pass_database

    bot.polling(none_stop=True, interval=0)

if (message.text == "\Update"):
    break;