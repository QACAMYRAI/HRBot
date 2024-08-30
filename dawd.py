import telebot
import time
from telebot import types
import openpyxl


TOKEN =

bot = telebot.TeleBot(TOKEN)

time_message = lambda x: time.strftime("%H:%M:%S %d.%m.%Y", time.localtime(x))              #Функция расчета времени
time_start_bot = time.strftime("%H:%M:%S %d.%m.%Y", time.localtime())

wb = openpyxl.load_workbook("HR.xlsx")                                            #Привязывает бота к таблице
sheet = wb.worksheets[0]

global user_name                                                                 #Создание глобальных переменных
user_name = ""

global user_about
user_about = ""

global flag
flag = 0

global user_city
user_city = ""

global user_ed
user_ed = ""

global user_exp
user_exp = ""

global user_game
user_game = ""


@bot.message_handler(commands=['start'])                                            #Приветственное сообщения бота в ответ на команду /start
def start(message):
    photo = open(r"start_photo.jpg", "rb")
    bot.send_photo(message.chat.id, photo)
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    TIER1 = types.KeyboardButton("Контакты")
    TIER2 = types.KeyboardButton("Ассортимент")
    TIER3 = types.KeyboardButton("Способ оплаты")
    TIER4 = types.KeyboardButton("HR")

    markup.add(TIER1, TIER2, TIER3, TIER4)  # Создание кнопок магазина

    bot.send_message(message.chat.id, "BYTEX – крупнейшая компания по тестированию игрового ПО в Восточной Европе и ведущий экспортёр Республики Мордовия в сфере IT. Компания образована в 2004 году в Саранске.  \n\n"
    "На сегодняшний день является экспертом на рынке тестирования, разработки и IT-образования. Среди партнёров компании - Mail.ru Group, Saber, Wargaming, 1C. \n\n"
    "BYTEX - это команда опытных профессионалов с исчерпывающими техническими возможностями. Штат компании насчитывает свыше 250 сотрудников, средний возраст которых – 27 лет. \n\n"
    "Нажмите на кнопку HR для заполнения анкеты", reply_markup=markup)
    global flag
    flag = 0

@bot.message_handler(content_types=['text'])
def bot_message(message):
    if message.chat.type == 'private' and time_start_bot <= time_message(message.date):                     #Условия для ответа бота
        if message.text == "Контакты":
            photo = open(r"karta.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            back = types.KeyboardButton("Назад")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(back, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Адрес: г. Саранск, ул. Советская, д.105а \n\n"
                                              "+7 (8342) 38-00-37 \n"
                                              "+7 (8342) 38-04-91 \n\n", reply_markup=markup)
            global flag
            flag = 0

        elif message.text == "Ассортимент":
            photo = open(r"assortiment.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            back = types.KeyboardButton("Назад")
            TIER1 = types.KeyboardButton("Ретро")
            TIER2 = types.KeyboardButton("Амбассадор")
            TIER3 = types.KeyboardButton("Геркулес")
            TIER4 = types.KeyboardButton("Спартанец")
            TIER5 = types.KeyboardButton("Бальбоа")
            markup.add(back, TIER1, TIER2, TIER3, TIER4, TIER5)
            bot.send_message(message.chat.id, "Нажмите на кнопку чтобы узнать информацию о костюме", reply_markup=markup)
            flag = 0

        elif message.text == "Ретро":
            photo = open(r"1.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            backtoassortyment = types.KeyboardButton("В ассортимент")
            markup.add(backtoassortyment)
            bot.send_message(message.chat.id, "Цена: 4900р \nСостав: Хлопок - 65%, Полиэстер - 35% \nРазмер товара на модели: L\nПараметры модели: 98-74-n\nАртикул MP002XM1HL6P", reply_markup=markup)
            flag = 0

        elif message.text == "Амбассадор":
            photo = open(r"2.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            backtoassortyment = types.KeyboardButton("В ассортимент")
            markup.add(backtoassortyment)
            bot.send_message(message.chat.id, "Цена: 4450р \nСостав: Хлопок - 65%, Полиэстер - 35% \nРазмер товара на модели: L\nПараметры модели: 98-74-n\nАртикул MP002XM1HL6P", reply_markup=markup)
            flag = 0

        elif message.text == "Геркулес":
            photo = open(r"3.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            backtoassortyment = types.KeyboardButton("В ассортимент")
            markup.add(backtoassortyment)
            bot.send_message(message.chat.id, "Цена: 5500р \nСостав: Хлопок - 65%, Полиэстер - 35% \nРазмер товара на модели: L\nПараметры модели: 98-74-n\nАртикул MP002XM1HL6P", reply_markup=markup)
            flag = 0

        elif message.text == "Спартанец":
            photo = open(r"4.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            backtoassortyment = types.KeyboardButton("В ассортимент")
            markup.add(backtoassortyment)
            bot.send_message(message.chat.id, "Цена: 4500р \nСостав: Хлопок - 65%, Полиэстер - 35% \nРазмер товара на модели: L\nПараметры модели: 98-74-n\nАртикул MP002XM1HL6P", reply_markup=markup)
            flag = 0

        elif message.text == "Бальбоа":
            photo = open(r"5.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            backtoassortyment = types.KeyboardButton("В ассортимент")
            markup.add(backtoassortyment)
            bot.send_message(message.chat.id, "Цена: 5200р \nСостав: Хлопок - 65%, Полиэстер - 35% \nРазмер товара на модели: L\nПараметры модели: 98-74-n\nАртикул MP002XM1HL6P", reply_markup=markup)
            flag = 0

        elif message.text == "Способ оплаты":
            photo = open(r"oplata.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            back = types.KeyboardButton("Назад")
            bot.send_message(message.chat.id, "Вы можете оплатить товар на нашем сайте http://gouspt.ru/ \n \nНа номер карты 2202 2002 3125 5752 \n \nИли в нашем магазине по адресу ул. Володарского, 20, Саранск", reply_markup=markup)
            flag = 0

        elif message.text == "В ассортимент":                                         #Кнопка назад в ассортимент
            photo = open(r"assortiment.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            back = types.KeyboardButton("Назад")
            TIER1 = types.KeyboardButton("Ретро")
            TIER2 = types.KeyboardButton("Амбассадор")
            TIER3 = types.KeyboardButton("Геркулес")
            TIER4 = types.KeyboardButton("Спартанец")
            TIER5 = types.KeyboardButton("Бальбоа")
            markup.add(back, TIER1, TIER2, TIER3, TIER4, TIER5)
            bot.send_message(message.chat.id, "Нажмите на кнопку чтобы узнать информацию о костюме", reply_markup=markup)
            flag = 0

        elif message.text == "HR":
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            bot.send_message(message.chat.id, "Введите ФИО", reply_markup=markup)
            flag = 1



        elif message.text == "Назад":                                         #Кнопка назад в начальное меню
            photo = open(r"start_photo.jpg", "rb")
            bot.send_photo(message.chat.id, photo)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")

            markup.add(TIER1, TIER2, TIER3, TIER4)

            bot.send_message(message.chat.id, "BYTEX – крупнейшая компания по тестированию игрового ПО в Восточной Европе и ведущий экспортёр Республики Мордовия в сфере IT. Компания образована в 2004 году в Саранске.  \n\n"
            "На сегодняшний день является экспертом на рынке тестирования, разработки и IT-образования. Среди партнёров компании - Mail.ru Group, Saber, Wargaming, 1C. \n\n"
            "BYTEX - это команда опытных профессионалов с исчерпывающими техническими возможностями. Штат компании насчитывает свыше 250 сотрудников, средний возраст которых – 27 лет. \n\n"
            "Нажмите на кнопку HR для заполнения анкеты", reply_markup=markup)

            flag = 0

        elif flag == 1:                                                                                       #Подтверждение ввода сообщения
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Да")
            TIER2 = types.KeyboardButton("Повторить ввод")
            markup.add(TIER1, TIER2)
            bot.send_message(message.chat.id, "Вы ввели" + " " + message.text, reply_markup=markup)
            flag = 2
            global user_name
            user_name = message.text


        elif flag == 2 and message.text == "Да":                                                               #Ввод данных в ячейки таблицы
            try:
                for i in range(1, 1000):
                    if sheet["A" + str(i)].value == None\
                            and sheet["B" + str(i)].value == None \
                            and sheet["C" + str(i)].value == None \
                            and sheet["D" + str(i)].value == None \
                            and sheet["E" + str(i)].value == None \
                            and sheet["F" + str(i)].value == None:
                        sheet["B" + str(i)].value = user_name.split(" ")[0]
                        sheet["C" + str(i)].value = user_name.split(" ")[1]
                        sheet["D" + str(i)].value = user_name.split(" ")[2]
                        iso = time.strftime('%Y-%m-%d %H:%M:%S')
                        sheet["A" + str(i)].value = iso
                        userid = str(message.from_user.username)
                        sheet["E" + str(i)].value = str('@' + userid)
                        break

            except IndexError:                                                                              #Повторная отправка вопроса при некорректных данных
                sheet["B" + str(i)].value = None
                sheet["C" + str(i)].value = None
                sheet["D" + str(i)].value = None
                markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                TIER1 = types.KeyboardButton("Контакты")
                TIER2 = types.KeyboardButton("Ассортимент")
                TIER3 = types.KeyboardButton("Способ оплаты")
                TIER4 = types.KeyboardButton("HR")

                markup.add(TIER1, TIER2, TIER3, TIER4)
                bot.send_message(message.chat.id, "Повторите ввод ФИО в верном формате", reply_markup=markup)
                flag = 1

            else:
                wb.save("HR.xlsx")
                markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                TIER1 = types.KeyboardButton("Контакты")
                TIER2 = types.KeyboardButton("Ассортимент")
                TIER3 = types.KeyboardButton("Способ оплаты")
                TIER4 = types.KeyboardButton("HR")
                markup.add(TIER1, TIER2, TIER3, TIER4)
                bot.send_message(message.chat.id, "Введите свой возраст", reply_markup=markup)
                flag = 3


        elif flag == 2 and message.text == "Повторить ввод":                                     #Подтверждение ввода сообщения
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Введите ФИО", reply_markup=markup)
            flag = 1

        elif flag == 3:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Да")
            TIER2 = types.KeyboardButton("Повторить ввод")
            markup.add(TIER1, TIER2)
            bot.send_message(message.chat.id, "Вы ввели" + " " + message.text, reply_markup=markup)
            flag = 4
            global user_about
            user_about = message.text

        elif flag == 4 and message.text == "Да":
            for i in range(1, 1000):
                if sheet["F" + str(i)].value == None:
                    sheet["F" + str(i)].value = user_about
                    break
            wb.save("HR.xlsx")
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Введите город проживания", reply_markup=markup)
            flag = 4

        elif flag == 4 and message.text == "Повторить ввод":
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Введите свой возраст", reply_markup=markup)
            flag = 3

        elif flag == 4:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Да")
            TIER2 = types.KeyboardButton("Повторить ввод")
            markup.add(TIER1, TIER2)
            bot.send_message(message.chat.id, "Вы ввели" + " " + message.text, reply_markup=markup)
            flag = 5
            global user_city
            user_city = message.text


        elif flag == 5 and message.text == "Да":
            for i in range(1, 1000):
                if sheet["G" + str(i)].value == None:
                    sheet["G" + str(i)].value = user_city
                    break
            wb.save("HR.xlsx")
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Укажите ваше образование", reply_markup=markup)
            flag = 5

        elif flag == 5 and message.text == "Повторить ввод":
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Введите город проживания", reply_markup=markup)
            flag = 4


        elif flag == 5:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Да")
            TIER2 = types.KeyboardButton("Повторить ввод")
            markup.add(TIER1, TIER2)
            bot.send_message(message.chat.id, "Вы ввели" + " " + message.text, reply_markup=markup)
            global user_ed
            user_ed = message.text
            flag = 6

        elif flag == 6 and message.text == "Да":
            for i in range(1, 1000):
                if sheet["H" + str(i)].value == None:
                    sheet["H" + str(i)].value = user_ed
                    break
            wb.save("HR.xlsx")
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Укажите ваш опыт работы", reply_markup=markup)
            flag = 6

        elif flag == 6 and message.text == "Повторить ввод":
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Укажите ваше образование", reply_markup=markup)
            flag = 5

        elif flag == 6:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Да")
            TIER2 = types.KeyboardButton("Повторить ввод")
            markup.add(TIER1, TIER2)
            bot.send_message(message.chat.id, "Вы ввели" + " " + message.text, reply_markup=markup)
            global user_exp
            user_exp = message.text
            flag = 7

        elif flag == 7 and message.text == "Да":
            for i in range(1, 1000):
                if sheet["I" + str(i)].value == None:
                    sheet["I" + str(i)].value = user_exp
                    break
            wb.save("HR.xlsx")
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "В какие игры в играли", reply_markup=markup)
            flag = 7

        elif flag == 7 and message.text == "Повторить ввод":
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Укажите ваш опыт работы", reply_markup=markup)
            flag = 6

        elif flag == 7:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Да")
            TIER2 = types.KeyboardButton("Повторить ввод")
            markup.add(TIER1, TIER2)
            bot.send_message(message.chat.id, "Вы ввели" + " " + message.text, reply_markup=markup)
            global user_game
            user_game = message.text
            flag = 8

        elif flag == 8 and message.text == "Да":
            for i in range(1, 1000):
                if sheet["J" + str(i)].value == None:
                    sheet["J" + str(i)].value = user_game
                    break
            wb.save("HR.xlsx")
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "Спасибо, ваша заявка принята", reply_markup=markup)
            flag = 7

        elif flag == 8 and message.text == "Повторить ввод":
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")
            markup.add(TIER1, TIER2, TIER3, TIER4)
            bot.send_message(message.chat.id, "В какие игры вы играли", reply_markup=markup)
            flag = 7



        else:                                                                #Блок отвечающий за отправку стикера в случае нераспознования команды
            badstiker = open(r"AnimatedSticker.tgs", "rb")
            bot.send_sticker(message.chat.id, badstiker)
            bot.send_message(message.chat.id, "Для удобства пользуйтесь кнопками, бот не распознает текст)")
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            TIER1 = types.KeyboardButton("Контакты")
            TIER2 = types.KeyboardButton("Ассортимент")
            TIER3 = types.KeyboardButton("Способ оплаты")
            TIER4 = types.KeyboardButton("HR")

            markup.add(TIER1, TIER2, TIER3, TIER4)

            bot.send_message(message.chat.id, reply_markup=markup)
            flag = 0


if __name__ == '__main__':                         #Блок отвечающий за перезапуск бота в случае падения
    while True:
        try:
            bot.polling(none_stop=True)
        except Exception as e:
            time.sleep(1)
            print(e)