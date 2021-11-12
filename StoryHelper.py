from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Updater, MessageHandler, Filters, CommandHandler
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




def main():
    updater = Updater(token=TOKEN)  # объект, который ловит сообщения от Telegrama

    dispather = updater.dispatcher

    start_help = CommandHandler('help', do_help)
    start_handler = CommandHandler('start', do_start)
    learn_handler = CommandHandler('learn', do_learn)
    changeclass_handler = MessageHandler(Filters.text, do_changeclass)
    some_handler = MessageHandler(Filters.text, do_some)
    echo = MessageHandler(Filters.all, do_echo)

    dispather.add_handler(learn_handler)
    dispather.add_handler(start_help)
    dispather.add_handler(start_handler)
    dispather.add_handler(changeclass_handler)
    dispather.add_handler(some_handler)
    dispather.add_handler(echo)

    updater.start_polling()
    updater.idle()


def do_echo(update, context):
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


def do_some(update: Update, context):
    text = update.message.text

    if text == 'Поиск Items':
        update.message.reply_text('', reply_markup=ReplyKeyboardRemove())
    elif text == 'Список недавно просмотренных Items':
        update.message.reply_text('Скоро...', reply_markup=ReplyKeyboardRemove())
    elif text == 'Отслеживание стоимости Items':
        update.message.reply_text('Скоро...', reply_markup=ReplyKeyboardRemove())
    elif text == 'Список ваших Items':
        update.message.reply_text('Скоро...', reply_markup=ReplyKeyboardRemove())
    else:
        update.message.reply_text('чаво', reply_markup=ReplyKeyboardRemove())


def do_learn(update, context):
    keyboard = [['7класс'], ['8класс'], ['9класс'],
                ['10класс'], ['11класс']]
    reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
    update.message.reply_text('Выберите класс.', reply_markup=reply_markup)


def do_changeclass(update: Update, context):
    text = update.message.text

    if text == '7класс':
        keyboard = [[sheet_7["A2"].value], ['тема2'], ['тема3'],
                    ['тема4'], ['тема5'], ['тема6']]
        reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
        update.message.reply_text('Выбирите тему, которую хотите изучать.', reply_markup=reply_markup)

    if text == f'{sheet_7["A2"].value}':
        keyboard = [[f'Материалы по "{sheet_7["A2"].value}"'], [f'Тесты по "{sheet_7["A2"].value}"']]
        reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
        update.message.reply_text('Хотите ознакомиться с материалами по теме или проверить свои знания?', reply_markup=reply_markup)

    if text == f'Материалы по "{sheet_7["A2"].value}"':
        keyboard = [[sheet_7["C2"].value], [sheet_7["C3"].value], [sheet_7["C4"].value]]
        reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
        update.message.reply_text('Выберите, то с чего хотите начать.', reply_markup=reply_markup)

    elif text == f'{sheet_7["C2"].value}':
        update.message.reply_text(f'{sheet_7["D2"].value}')
    elif text == f'{sheet_7["C3"].value}':
        update.message.reply_text(f'{sheet_7["D3"].value}')
    elif text == f'{sheet_7["C4"].value}':
        update.message.reply_text(f'{sheet_7["D4"].value}')

    if text == f'Тесты по "{sheet_7["A2"].value}"':
        keyboard = [[sheet_7["E2"].value], [sheet_7["E3"].value], [sheet_7["E4"].value]]
        reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, one_time_keyboard=True, resize_keyboard=True)
        update.message.reply_text('Выбирите тест.', reply_markup=reply_markup)

    elif text == f'{sheet_7["E2"].value}':
        update.message.reply_text(f'{sheet_7["F2"].value}')
    elif text == f'{sheet_7["E3"].value}':
        update.message.reply_text(f'{sheet_7["F3"].value}')
    elif text == f'{sheet_7["E4"].value}':
        update.message.reply_text(f'{sheet_7["F4"].value}')


main()