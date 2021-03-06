from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Updater, MessageHandler, Filters, CommandHandler, CallbackContext
from openpyxl import load_workbook

TOKEN = '2132309010:AAFQMIsoBEGKZYIR5kKy__dCI9keXwCIWKw'
book = load_workbook("DATAofStoryBotHelper.xlsx")

sheet_1 = book['Лист1']
sheet_2 = book['Лист2']
sheet_3 = book['Лист3']
sheet_4 = book['Лист4']
sheet_5 = book['Лист5']
sheet_6 = book['Лист6']
sheet_7 = book['Лист7']
sheet_8 = book['Лист8']
sheet_9 = book['Лист9']
sheet_10 = book['Лист10']
sheet_11 = book['Лист11']

COLUMN_THEME = 1
COLUMN_MATERIAL = 3
COLUMN_MATERIAL_LINK = 4
COLUMN_TEST = 5
COLUMN_TEST_LINK = 6

page = book[]


def main():
    updater = Updater(token=TOKEN)  # объект, который ловит сообщения от Telegrama

    dispather = updater.dispatcher

    start_help = CommandHandler('help', do_help)
    start_handler = CommandHandler('start', do_start)
    learn_handler = CommandHandler('learn', do_learn)
    cheker_handler = CommandHandler('check', do_cheker)
    changeclass_handler = MessageHandler(Filters.text, do_changeclass)
    echo = MessageHandler(Filters.all, do_echo)

    dispather.add_handler(learn_handler)
    dispather.add_handler(start_help)
    dispather.add_handler(start_handler)
    dispather.add_handler(cheker_handler)
    dispather.add_handler(changeclass_handler)

    dispather.add_handler(echo)

    updater.start_polling()
    updater.idle()


def do_echo(update: Update, context):
    update.message.reply_text('ЧаВо?')




def do_start(update, context):
    update.message.reply_text('Здавствуйте! Я бот-Историк. Я помогу вам изучать историю,'
                              '\nа так же помогу закрепить пройденное.'
                              '\nЧтобы ознакомится с функционалом, напишите - /help.'
                              )


def do_help(update, context):
    update.message.reply_text('Даный бот поможет вам самостоятельно изучать темы по истории и проходить тесты по ним.'
                              '\nЧтобы выбрать тему напишите /learn'
                              '\nЧто делать дальше?'
                              '\n1. Выбирите свой класс.'
                              '\n2. Выбирите тему, которая вам нужна.'
                              '\n3. Выбирите то, что вам нужно. Материалы по теме или тесты по теме.'
                              '\n4. Изучайте темы и проходите тесты. Удачи в обучении!'
                              '\nНашли баг или опечатку, или же что-то не работает? По вопросам:(тут будет почта)'
                             )


def do_cheker(update,context):
    keyboard = [['ячейки_класс7'], ['ячейки_класс8'], ['ячейки_класс9'],
                ['ячейки_класс10'], ['ячейки_класс11']]
    reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
    update.message.reply_text('Выберите класс.', reply_markup=reply_markup)


def do_learn(update, context):
    keyboard = [['7класс'], ['8класс'], ['9класс'],
                ['10класс'], ['11класс']]
    reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
    update.message.reply_text('Выберите класс.', reply_markup=reply_markup)


def do_changeclass(update: Update, context):
    text = update.message.text

    if text == 'ячейки_класс7':
        update.message.reply_text(f'Список доступных ячеек'
                              f'\nТема -A2:{sheet_7["A2"].value}'
                              f'\nМатериалы(ссылки) -C2:{sheet_7["C2"].value} D2:{sheet_7["D2"].value}'
                              f'\nМатериалы(ссылки) -C3:{sheet_7["C2"].value} D3:{sheet_7["D3"].value}'
                              f'\nМатериалы(ссылки) -C4:{sheet_7["C2"].value} D4:{sheet_7["D4"].value}'
                              f'\nТесты - E2:{sheet_7["E2"].value} F2:{sheet_7["F2"].value}'
                              f'\nТесты - E3:{sheet_7["E3"].value} F3:{sheet_7["F3"].value}'
                              f'\nТесты - E4:{sheet_7["E4"].value} F4:{sheet_7["F4"].value}', reply_markup=ReplyKeyboardRemove())

    #if text == '7класс':
     #   keyboard = [[sheet_7["A2"].value], ['тема2'], ['тема3'],
     #               ['тема4'], ['тема5'], ['тема6']]
   #     reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
    #    update.message.reply_text('Выбирите тему, которую хотите изучать.', reply_markup=reply_markup)

   # if text == f'{sheet_7["A2"].value}':
   #     keyboard = [[f'Материалы по "{sheet_7["A2"].value}"'], [f'Тесты по "{sheet_7["A2"].value}"']]
   #     reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
   #     update.message.reply_text('Хотите ознакомиться с материалами по теме или проверить свои знания?', reply_markup=reply_markup)

    #if text == f'Материалы по "{sheet_7["A2"].value}"':
      #  keyboard = [[sheet_7["C2"].value], [sheet_7["C3"].value], [sheet_7["C4"].value]]
      #  reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
     #   update.message.reply_text('Выберите, то с чего хотите начать.', reply_markup=reply_markup)

    #elif text == f'{sheet_7["C2"].value}':
     #   update.message.reply_text(f'{sheet_7["D2"].value}', reply_markup=ReplyKeyboardRemove())
    #elif text == f'{sheet_7["C3"].value}':
     #   update.message.reply_text(f'{sheet_7["D3"].value}', reply_markup=ReplyKeyboardRemove())
    #elif text == f'{sheet_7["C4"].value}':
       # update.message.reply_text(f'{sheet_7["D4"].value}', reply_markup=ReplyKeyboardRemove())

   # if text == f'Тесты по "{sheet_7["A2"].value}"':
    #    keyboard = [[sheet_7["E2"].value], [sheet_7["E3"].value], [sheet_7["E4"].value]]
     #   reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
      #  update.message.reply_text('Выбирите тест.', reply_markup=reply_markup)

  #  elif text == f'{sheet_7["E2"].value}':
 #       update.message.reply_text(f'{sheet_7["F2"].value}', reply_markup=ReplyKeyboardRemove())
 #   elif text == f'{sheet_7["E3"].value}':
#        update.message.reply_text(f'{sheet_7["F3"].value}', reply_markup=ReplyKeyboardRemove())
 #  elif text == f'{sheet_7["E4"].value}':
#        update.message.reply_text(f'{sheet_7["F4"].value}', reply_markup=ReplyKeyboardRemove())

    if text == 'Материалы':
        return do_material_check(update, context)

    if text == 'Тесты':
        return do_test_check(update, context)

    if text == '/check_value':
        return do_learn(update, context)


def do_theme_check(update, context): #проверяет существующие темы
    theme = []
    for line in range(2, sheet_7.max_row + 1):
        theme.append(sheet_7.cell(row=line, column=COLUMN_THEME).value)
    theme = set(theme)
    keyboard = [theme]
    key_board = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
    update.message.reply_text('Посмотрите', reply_markup=key_board)
    answer = update.message.text
    return do_quest(update, context)


def do_quest(update, context): # материалы или тесты
    keyboard = [['Материалы'], ['Тесты']]
    key_board = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
    update.message.reply_text('Посмотрите', reply_markup=key_board)


def do_material_check(update, context): # смотрит все материалы для заданной темы
    material = []

    for line in range(2, sheet_7.max_row + 1):
        value = sheet_7.cell(row=line, column=COLUMN_THEME).value
        if value == answer:
            material.append(sheet_7.cell(row=line, column=COLUMN_MATERIAL).value)
            keyboard = [material]
            key_board = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
            update.message.reply_text('Посмотрите', reply_markup=key_board)


def do_test_check(update, contex): # смотрит все тесты для заданной темы
    test = []

    for line in range(2, sheet_7.max_row + 1):
        value = sheet_7.cell(row=line, column=COLUMN_THEME).value
        if value == answer:
            test.append(sheet_7.cell(row=line, column=COLUMN_TEST).value)
            keyboard = [test]
            key_board = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
            update.message.reply_text('Посмотрите', reply_markup=key_board)


#def do_set_theme_name(update, context):

#def do_set_material_name(update, context):

#def do_set_test_name(update, context):

#def do_set_material_link(update, context):

#def do_set_test_link(update, context):


main()
